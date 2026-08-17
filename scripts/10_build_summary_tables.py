"""
10_build_summary_tables.py
Build the paper's summary tables — README tab + 12 static summary tabs.
Written to Summary_tables_v6.xlsx and copied to Annex_C.xlsx (the committed
paper deliverable).

v2 changes vs v1:
  - Tab order within each group reversed: GHG incl → GHG excl → CO2 incl → CO2 excl
  - World rows swapped: "World (all emissions)" first, "World (country-attributable)" second
  - Blank spacer row added after each group section
  - Colonial tabs now include same transport delta as regular 1850 tabs
  - CGD official colours applied (dark teal / medium teal / amber)
  - Font updated to Sofia Pro throughout
  - README tab added (first tab): table of contents + source coverage + notes

12 data tabs = 4 measures × 3 period/type combinations:
  Tabs 1–4:   1850–2024, as-is
  Tabs 5–8:   1990–2024, as-is
  Tabs 9–12:  1850–2024, colonial-attributed

Row structure (same across all data tabs):
  World (all emissions)          — GtCO2e  [incl. int'l transport where available]
  World (country-attributable)   — GtCO2e  [sum of country totals only]
  [spacer]
  By institutional group: Annex I / Annex II / Annex II+EU / Non-Annex I /
                          OECD / EU / BRICS / G7 / G20 / Non-Annex I G20 members
  [spacer]
  By income group: HIC / UMIC / LMIC / LIC
  [spacer]
  By development group: LDC / LLDC / SIDS / LIC+LMIC+SIDS+LDC
  [spacer]
  By region (WB): 7 regions
  [spacer]
  G20 countries: all 19 G20 member states (alphabetical)

Sources per tab:
  1850–2024: OWID, PRIMAP + GCP_fossil (CO2 excl) or GCP_BLUE/OSCAR/LUCE (CO2 incl)
  1990–2024: all of above + EDGAR + Climate Watch
  Colonial:  same as 1850–2024 (OWID, PRIMAP, GCP variants)

World (all emissions) approach per source:
  OWID         — World entity in raw CSV (includes aviation + shipping)
  PRIMAP       — same as country-attributable: no int'l transport in dataset
  EDGAR        — country sum + AIR + SEA entities
  ClimateWatch — WORLD entity in raw CSV (includes bunker fuels)
  GCP fossil   — World column in raw xlsx (country sum + shipping + aviation)
  GCP LUC      — same as country-attributable: no transport in land-use models

Colonial computation:
  countries_annual × colonial_attribution_long joined on iso_code + year.
  Countries absent from colonial database attributed to themselves (coefficient=1).
  non_annex_1 derived as World − annex_1 (no direct flag in country_groups.csv).
  World (all emissions) for colonial tabs = same as regular 1850–2024 tabs
  (transport is a global total added on top; colonial attribution does not affect it).
"""

import os
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Paths ──────────────────────────────────────────────────────────────────────
SCRIPTS_DIR     = os.path.dirname(os.path.abspath(__file__))
QA_DIR          = os.path.dirname(SCRIPTS_DIR)
RAW_DIR         = os.path.join(QA_DIR, "data", "raw")
OUT_DIR         = os.path.join(QA_DIR, "data", "outputs")

GROUPS_CSV      = os.path.join(OUT_DIR, "groups_summary.csv")
COUNTRIES_CSV   = os.path.join(OUT_DIR, "countries_annual.csv")
COLONIAL_CSV    = os.path.join(OUT_DIR, "colonial_attribution_long.csv")
CG_CSV          = os.path.join(QA_DIR, "data", "country_groups.csv")

OWID_RAW        = os.path.join(RAW_DIR, "owid-co2-data.csv")
EDGAR_XLSX_PATH = os.path.join(RAW_DIR, "EDGAR_2025_GHG_booklet_2025.xlsx")
EDGAR_FALLBACK  = "c:/tmp/edgar_booklet_tmp.xlsx"
CW_RAW          = os.path.join(RAW_DIR, "CW_HistoricalEmissions_ClimateWatch.csv")
GCP_FOSSIL_XLSX = os.path.join(RAW_DIR, "National_Fossil_Carbon_Emissions_2025_v0.3.xlsx")

OUT_XLSX        = os.path.join(OUT_DIR, "Summary_tables_v6.xlsx")
ANNEX_C_XLSX    = os.path.join(OUT_DIR, "Annex_C.xlsx")   # paper deliverable copy

GCP_CONV = 3.664   # MtC → MtCO2

# ── Measures: GHG incl → GHG excl → CO2 incl → CO2 excl ─────────────────────
MEASURES = [
    ("ghg_incLUC", "GHG inc LULUCF"),
    ("ghg_excLUC", "GHG exc LULUCF"),
    ("co2_incLUC", "CO2 inc LULUCF"),
    ("co2_excLUC", "CO2 exc LULUCF"),
]

# ── 12 tab definitions + README ────────────────────────────────────────────────
# (tab_name, start_year, end_year, colonial, measure_code)
TABS = []
for m_code, m_label in MEASURES:
    TABS.append((f"{m_label} - 1850",     1850, 2024, False, m_code))
for m_code, m_label in MEASURES:
    TABS.append((f"{m_label} - 1990",     1990, 2024, False, m_code))
for m_code, m_label in MEASURES:
    TABS.append((f"{m_label} - colonial", 1850, 2024, True,  m_code))

# ── Source labels ──────────────────────────────────────────────────────────────
SOURCE_LABELS = {
    "OWID":         "OWID (GCB 2025)",
    "PRIMAP":       "PRIMAP-hist v2.7",
    "EDGAR":        "EDGAR 2025",
    "ClimateWatch": "Climate Watch",
    "GCP_fossil":   "GCP fossil",
    "GCP_BLUE":     "GCP BLUE",
    "GCP_OSCAR":    "GCP OSCAR",
    "GCP_LUCE":     "GCP LUCE",
    "GMST":         "GMST (°C / %)",
}

# Source metadata for README coverage table
SOURCE_META = [
    ("OWID",         "OWID (GCB 2025)",    "1750–2024 (CO2); 1850–2024 (GHGs)",
     "All 4",        "Aviation + shipping (OWID World entity)"),
    ("PRIMAP",       "PRIMAP-hist v2.7",   "1750–2024",
     "All 4",        "None — World (all) = World (country-attr.) for this source"),
    ("EDGAR",        "EDGAR 2025",         "1970–2024 (excl LULUCF); 1990–2024 (incl LULUCF)",
     "All 4",        "AIR + SEA entities from EDGAR booklet"),
    ("ClimateWatch", "Climate Watch",      "1990–2023",
     "All 4",        "Bunker fuels (WORLD entity; includes int'l maritime + aviation bunkers)"),
    ("GCP_fossil",   "GCP fossil",         "1750–2024",
     "CO2 excl LULUCF only",
     "International shipping + aviation (GCP 'World' column in raw xlsx)"),
    ("GCP_BLUE",     "GCP BLUE",           "1750–2024",
     "CO2 incl LULUCF only",
     "GCP fossil international transport (bunkers) added to country-sum LUC+fossil"),
    ("GCP_OSCAR",    "GCP OSCAR",          "1750–2024",
     "CO2 incl LULUCF only",
     "GCP fossil international transport (bunkers) added to country-sum LUC+fossil"),
    ("GCP_LUCE",     "GCP LUCE",           "1750–2024",
     "CO2 incl LULUCF only",
     "GCP fossil international transport (bunkers) added to country-sum LUC+fossil"),
    ("GMST",         "GMST (Jones et al.)", "1851–2024",
     "All 4 (gmst_ prefix)",
     "No transport — GMST World rows = country sum (°C). Colonial tabs show as-is GMST."),
]


def tab_sources(start_yr, measure):
    """Ordered list of (source_code, source_label) for a tab. GMST always last."""
    srcs = ["OWID", "PRIMAP"]
    if start_yr >= 1990:
        srcs += ["EDGAR", "ClimateWatch"]
    if measure == "co2_excLUC":
        srcs.append("GCP_fossil")
    elif measure == "co2_incLUC":
        srcs += ["GCP_BLUE", "GCP_OSCAR", "GCP_LUCE"]
    srcs.append("GMST")
    return [(s, SOURCE_LABELS[s]) for s in srcs]


# ── Row definitions ────────────────────────────────────────────────────────────
# world_all FIRST, world_sum SECOND; spacer entries between sections
# Non-Annex I G20 members = G20 members not in Annex I (replaced Gulf states 2026-07-07)
NAI_G20_ISOS = ["ARG", "BRA", "CHN", "IDN", "IND", "KOR", "MEX", "SAU", "ZAF"]

ROW_DEFS = [
    # (type,        code,                label)
    ("world_all",  None,                "World (all emissions)"),
    ("world_sum",  None,                "World (country-attributable)"),
    ("spacer",     None,                ""),
    ("section",    None,                "By institutional group"),
    ("group",      "annex_1",           "Annex I"),
    ("group",      "annex_2",           "Annex II"),
    ("group",      "annex_2_eu",        "Annex II + other EU Members"),
    ("group",      "non_annex_1",       "Non-Annex I"),
    ("group",      "oecd",              "OECD"),
    ("group",      "eu",                "EU"),
    ("group",      "brics",             "BRICS"),
    ("group",      "g7",                "G7"),
    ("group",      "g20",               "G20"),
    ("custom",     "NAI_G20",           "Non-Annex I G20 members"),
    ("spacer",     None,                ""),
    ("section",    None,                "By income group"),
    ("group",      "hic",               "HIC"),
    ("group",      "umic",              "UMIC"),
    ("group",      "lmic",              "LMIC"),
    ("group",      "lic",               "LIC"),
    ("spacer",     None,                ""),
    ("section",    None,                "By development group"),
    ("group",      "ldc",               "LDC"),
    ("group",      "lldc",              "LLDC"),
    ("group",      "sids",              "SIDS"),
    ("group",      "lic_lmic_sids_ldc", "LIC / LMIC / SIDS / LDC"),
    ("spacer",     None,                ""),
    ("section",    None,                "By region"),
    ("group",      "wb_eca",            "Europe and Central Asia"),
    ("group",      "wb_n_am",           "North America"),
    ("group",      "wb_lac",            "Latin America & Caribbean"),
    ("group",      "wb_eap",            "East Asia & Pacific"),
    ("group",      "wb_sa",             "South Asia"),
    ("group",      "wb_mena",           "Middle East & North Africa"),
    ("group",      "wb_ssa",            "Sub-Saharan Africa"),
    ("spacer",     None,                ""),
    ("section",    None,                "G20 countries"),
    ("country",    "ARG",               "Argentina"),
    ("country",    "AUS",               "Australia"),
    ("country",    "BRA",               "Brazil"),
    ("country",    "CAN",               "Canada"),
    ("country",    "CHN",               "China"),
    ("country",    "FRA",               "France"),
    ("country",    "DEU",               "Germany"),
    ("country",    "IND",               "India"),
    ("country",    "IDN",               "Indonesia"),
    ("country",    "ITA",               "Italy"),
    ("country",    "JPN",               "Japan"),
    ("country",    "MEX",               "Mexico"),
    ("country",    "RUS",               "Russia"),
    ("country",    "SAU",               "Saudi Arabia"),
    ("country",    "ZAF",               "South Africa"),
    ("country",    "KOR",               "South Korea"),
    ("country",    "TUR",               "Turkey"),
    ("country",    "GBR",               "United Kingdom"),
    ("country",    "USA",               "United States"),
]

# ── CGD colours ────────────────────────────────────────────────────────────────
CGD_DARK_TEAL   = "0B4C5B"   # primary — 1850 tab headers
CGD_MED_TEAL    = "1A8A9E"   # derived  — 1990 tab headers
CGD_AMBER       = "FFB52C"   # accent   — colonial tab headers
CGD_PALE_TEAL   = "A8D5DB"   # UI pale teal — 1990 section rows
CGD_NW_TEAL     = "D6EAED"   # near-white teal — 1850 section rows
CGD_PALE_AMBER  = "FFF5D9"   # derived pale amber — colonial section rows
GREY_LIGHT      = "F2F2F2"   # world rows
WHITE           = "FFFFFF"
TEXT_DARK       = "0B4C5B"   # near-black replacement → dark teal for headers on light bg
TEXT_MUTED      = "777777"
FONT_HEAD       = "Sofia Pro Semi Bold"
FONT_BODY       = "Sofia Pro"

# Per colour-group: (header_fill, header_text, section_fill, section_text, tab_colour)
BLACK = "000000"

COLOUR_SCHEME = {
    "1850":     (CGD_DARK_TEAL, WHITE,  CGD_NW_TEAL,    TEXT_DARK, CGD_DARK_TEAL),
    "1990":     (CGD_MED_TEAL,  WHITE,  CGD_PALE_TEAL,  TEXT_DARK, CGD_MED_TEAL),
    "colonial": (CGD_AMBER,     BLACK,  CGD_PALE_AMBER, BLACK,     CGD_AMBER),
}


def colour_key(start_yr, is_colonial):
    if is_colonial:
        return "colonial"
    return "1850" if start_yr == 1850 else "1990"


def _f(bold=False, italic=False, color="000000", size=10, font_name=None):
    name = font_name or FONT_BODY
    return Font(bold=bold, italic=italic, color=color, name=name, size=size)


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


ALIGN_C  = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_R  = Alignment(horizontal="right",  vertical="center")
ALIGN_L  = Alignment(horizontal="left",   vertical="center")
ALIGN_LW = Alignment(horizontal="left",   vertical="center", wrap_text=True)

EM_DASH = "—"


# ═══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════════════════════

print("Loading pipeline data ...")
groups   = pd.read_csv(GROUPS_CSV)
ca       = pd.read_csv(COUNTRIES_CSV)
colonial = pd.read_csv(COLONIAL_CSV)
cg       = pd.read_csv(CG_CSV)

print(f"  groups_summary: {len(groups):,} rows")
print(f"  countries_annual: {len(ca):,} rows")
print(f"  colonial_attribution_long: {len(colonial):,} rows")

# Pre-index GMST data for fast point lookups (cumulative °C at each year)
print("  Indexing GMST data ...")
_gmst_grp = (groups[groups["source"] == "GMST"]
             .set_index(["group", "measure", "year"])["value_Mt"])
_gmst_ctry = (ca[ca["source"] == "GMST"]
              .set_index(["iso_code", "measure", "year"])["value_Mt"])


# ── Cumulation helpers ─────────────────────────────────────────────────────────

def cum_group(grp_code, source, measure, s, e):
    mask = (
        (groups["group"]   == grp_code) &
        (groups["source"]  == source) &
        (groups["measure"] == measure) &
        (groups["year"]    >= s) &
        (groups["year"]    <= e)
    )
    sub = groups.loc[mask, "value_Mt"]
    return np.nan if sub.empty else float(sub.sum())


def cum_country(iso, source, measure, s, e):
    mask = (
        (ca["iso_code"] == iso) &
        (ca["source"]   == source) &
        (ca["measure"]  == measure) &
        (ca["year"]     >= s) &
        (ca["year"]     <= e)
    )
    sub = ca.loc[mask, "value_Mt"]
    return np.nan if sub.empty else float(sub.sum())


def cum_nai_g20(source, measure, s, e):
    vals = [v for iso in NAI_G20_ISOS
            if not np.isnan(v := cum_country(iso, source, measure, s, e))]
    return float(sum(vals)) if vals else np.nan


# ═══════════════════════════════════════════════════════════════════════════════
# WORLD (ALL EMISSIONS) EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════

def _edgar_path():
    try:
        open(EDGAR_XLSX_PATH, "rb").close()
        return EDGAR_XLSX_PATH
    except PermissionError:
        return EDGAR_FALLBACK


def _load_owid_world():
    print("  Loading OWID World entity ...")
    df = pd.read_csv(OWID_RAW, usecols=[
        "country", "year", "co2", "land_use_change_co2",
        "total_ghg", "total_ghg_excluding_lucf"
    ])
    world = df[df["country"] == "World"].copy()
    both_nan = world["co2"].isna() & world["land_use_change_co2"].isna()
    world["co2_incLUC"] = np.where(
        both_nan, np.nan,
        world["co2"].fillna(0) + world["land_use_change_co2"].fillna(0)
    )
    world = world.rename(columns={
        "co2":                      "co2_excLUC",
        "total_ghg_excluding_lucf": "ghg_excLUC",
        "total_ghg":                "ghg_incLUC",
    })
    return world[["year", "co2_excLUC", "co2_incLUC", "ghg_excLUC", "ghg_incLUC"]].copy()


def _load_edgar_transport():
    print("  Loading EDGAR transport (AIR+SEA) ...")
    try:
        df = pd.read_excel(_edgar_path(), sheet_name="GHG_by_sector_and_country")
    except Exception as exc:
        print(f"    WARNING: {exc}")
        return None, None
    df.columns = [str(c).strip() for c in df.columns]
    year_cols = [c for c in df.columns if str(c).isdigit() and 1900 <= int(c) <= 2100]
    iso_col   = "EDGAR Country Code"
    transport = df[df[iso_col].isin({"AIR", "SEA"})].copy()
    long = transport.melt(
        id_vars=[iso_col, "Substance"],
        value_vars=year_cols, var_name="year", value_name="value_Mt"
    )
    long["year"]     = long["year"].astype(int)
    long["value_Mt"] = pd.to_numeric(long["value_Mt"], errors="coerce")
    long = long.dropna(subset=["value_Mt"])
    co2_t = long[long["Substance"] == "CO2"].groupby("year")["value_Mt"].sum()
    ghg_t = long.groupby("year")["value_Mt"].sum()
    print(f"    CO2 transport: {len(co2_t)} year rows; GHG transport: {len(ghg_t)} year rows")
    return co2_t, ghg_t


def _load_cw_world():
    print("  Loading ClimateWatch WORLD entity ...")
    df = pd.read_csv(CW_RAW, low_memory=False)
    df["Sector"] = df["Sector"].str.strip()
    year_cols = [c for c in df.columns if str(c).isdigit()]
    MEASURE_MAP = {
        ("CO2",     "Total excluding LULUCF"): "co2_excLUC",
        ("CO2",     "Total including LULUCF"): "co2_incLUC",
        ("All GHG", "Total excluding LULUCF"): "ghg_excLUC",
        ("All GHG", "Total including LULUCF"): "ghg_incLUC",
    }
    result = {}
    for (gas, sector), measure in MEASURE_MAP.items():
        row = df[
            (df["Country"] == "WORLD") &
            (df["Gas"]     == gas) &
            (df["Sector"]  == sector)
        ]
        if row.empty:
            print(f"    WARNING: no CW WORLD row for {gas} / {sector}")
            continue
        long = row.melt(id_vars=["Country"], value_vars=year_cols,
                        var_name="year", value_name="value_Mt")
        long["year"] = long["year"].astype(int)
        long = long.dropna(subset=["value_Mt"])
        result[measure] = long.set_index("year")["value_Mt"]
    return result


def _load_gcp_world_all():
    """GCP fossil 'World' column = country sum + int'l shipping + aviation."""
    print("  Loading GCP fossil World total ...")
    wb = openpyxl.load_workbook(GCP_FOSSIL_XLSX, read_only=True, data_only=True)
    ws = wb["Territorial Emissions"]
    world_col = None
    for cell in next(ws.iter_rows(min_row=11, max_row=11)):
        if cell.value is not None and str(cell.value).strip() == "World":
            world_col = cell.column
            break
    wb.close()
    if world_col is None:
        print("    WARNING: 'World' column not found in GCP fossil file")
        return pd.Series(dtype=float)
    wb = openpyxl.load_workbook(GCP_FOSSIL_XLSX, read_only=True, data_only=True)
    ws = wb["Territorial Emissions"]
    rows = []
    for r in ws.iter_rows(min_row=13, values_only=True):
        yr = r[0]
        if not isinstance(yr, (int, float)):
            continue
        val = r[world_col - 1]
        if val is not None:
            rows.append({"year": int(yr), "value_Mt": float(val) * GCP_CONV})
    wb.close()
    if not rows:
        return pd.Series(dtype=float)
    s = pd.DataFrame(rows).set_index("year")["value_Mt"]
    print(f"    GCP World total: {len(s)} year rows, latest = {s.iloc[-1]:.0f} MtCO2")
    return s


print("Loading World (all emissions) data ...")
OWID_WORLD    = _load_owid_world()
EDGAR_CO2_T, EDGAR_GHG_T = _load_edgar_transport()
CW_WORLD      = _load_cw_world()
GCP_WORLD_ALL = _load_gcp_world_all()


def _gmst_at_grp(grp, gmst_m, yr):
    """Cumulative GMST (°C) for a group at a single year; 0.0 if yr < 1851."""
    if yr < 1851:
        return 0.0
    try:
        return float(_gmst_grp.loc[(grp, gmst_m, yr)])
    except KeyError:
        return np.nan


def _gmst_at_ctry(iso, gmst_m, yr):
    """Cumulative GMST (°C) for a country at a single year; 0.0 if yr < 1851."""
    if yr < 1851:
        return 0.0
    try:
        return float(_gmst_ctry.loc[(iso, gmst_m, yr)])
    except KeyError:
        return np.nan


def gmst_delta_group(grp_code, measure, s, e):
    """GMST contribution (°C) from period [s, e] for a group."""
    gmst_m = "gmst_" + measure
    ve = _gmst_at_grp(grp_code, gmst_m, e)
    vs = _gmst_at_grp(grp_code, gmst_m, s - 1)
    if np.isnan(ve):
        return np.nan
    return ve - (0.0 if np.isnan(vs) else vs)


def gmst_delta_country(iso, measure, s, e):
    """GMST contribution (°C) from period [s, e] for a country."""
    gmst_m = "gmst_" + measure
    ve = _gmst_at_ctry(iso, gmst_m, e)
    vs = _gmst_at_ctry(iso, gmst_m, s - 1)
    if np.isnan(ve):
        return np.nan
    return ve - (0.0 if np.isnan(vs) else vs)


def gmst_delta_nai_g20(measure, s, e):
    vals = [v for iso in NAI_G20_ISOS
            if not np.isnan(v := gmst_delta_country(iso, measure, s, e))]
    return float(sum(vals)) if vals else np.nan


def world_all_Mt(source, measure, s, e):
    """World (all emissions) total Mt for source × measure × year range.
    Used for BOTH regular and colonial tabs."""
    if source == "OWID":
        if measure == "co2_incLUC":
            # OWID's World entity LUC != sum of country LUC (different GCP global
            # estimate vs national totals). Use country sum + fossil transport delta.
            world_sum = cum_group("World", "OWID", "co2_incLUC", s, e)
            if np.isnan(world_sum):
                return np.nan
            owid_co2_world = float(OWID_WORLD[
                (OWID_WORLD["year"] >= s) & (OWID_WORLD["year"] <= e)
            ]["co2_excLUC"].sum())
            country_co2 = cum_group("World", "OWID", "co2_excLUC", s, e)
            transport = owid_co2_world - (country_co2 if not np.isnan(country_co2) else owid_co2_world)
            return float(world_sum + transport)
        sub = OWID_WORLD[
            (OWID_WORLD["year"] >= s) & (OWID_WORLD["year"] <= e)
        ][measure]
        return float(sub.sum()) if not sub.dropna().empty else np.nan

    elif source == "PRIMAP":
        return cum_group("World", "PRIMAP", measure, s, e)

    elif source == "EDGAR":
        world_sum = cum_group("World", "EDGAR", measure, s, e)
        if np.isnan(world_sum):
            return np.nan
        if EDGAR_CO2_T is None:
            return world_sum
        t_series = EDGAR_CO2_T if measure in ("co2_excLUC", "co2_incLUC") else EDGAR_GHG_T
        t_sub = t_series[(t_series.index >= s) & (t_series.index <= e)]
        delta = float(t_sub.sum()) if not t_sub.empty else 0.0
        return world_sum + delta

    elif source == "ClimateWatch":
        if measure not in CW_WORLD:
            return np.nan
        series = CW_WORLD[measure]
        sub = series[(series.index >= s) & (series.index <= e)]
        return float(sub.sum()) if not sub.empty else np.nan

    elif source == "GCP_fossil":
        if measure != "co2_excLUC":
            return cum_group("World", "GCP_fossil", measure, s, e)
        if GCP_WORLD_ALL.empty:
            return cum_group("World", "GCP_fossil", measure, s, e)
        sub = GCP_WORLD_ALL[(GCP_WORLD_ALL.index >= s) & (GCP_WORLD_ALL.index <= e)]
        return float(sub.sum()) if not sub.empty else np.nan

    elif source in ("GCP_BLUE", "GCP_OSCAR", "GCP_LUCE"):
        world_sum = cum_group("World", source, measure, s, e)
        if np.isnan(world_sum):
            return np.nan
        # Add GCP fossil transport delta (bunkers excluded from country totals)
        if not GCP_WORLD_ALL.empty:
            sub_all = GCP_WORLD_ALL[(GCP_WORLD_ALL.index >= s) & (GCP_WORLD_ALL.index <= e)]
            gcp_fossil_sum = cum_group("World", "GCP_fossil", "co2_excLUC", s, e)
            if not sub_all.empty and not np.isnan(gcp_fossil_sum):
                return float(world_sum + float(sub_all.sum()) - gcp_fossil_sum)
        return world_sum

    return np.nan


# ═══════════════════════════════════════════════════════════════════════════════
# COLONIAL ATTRIBUTION
# ═══════════════════════════════════════════════════════════════════════════════

CG_GROUP_COLS = [c for c in cg.columns if c not in ("iso_code", "country_name")]
_colonial_cache = {}


def compute_colonial(measure, s, e):
    """Returns {(source, kind, code): value_Mt}.
    kind = 'world' | 'group' | 'country'
    """
    key = (measure, s, e)
    if key in _colonial_cache:
        return _colonial_cache[key]

    print(f"  Computing colonial attribution: {measure} {s}–{e} ...")

    ca_sub = ca[
        (ca["measure"] == measure) &
        (ca["year"]    >= s) &
        (ca["year"]    <= e)
    ][["iso_code", "source", "year", "value_Mt"]].copy()

    if ca_sub.empty:
        _colonial_cache[key] = {}
        return {}

    col_sub = colonial[
        (colonial["year"] >= s) & (colonial["year"] <= e)
    ][["territory_iso", "year", "attributed_to_iso", "coefficient"]].copy()

    merged = ca_sub.merge(
        col_sub,
        left_on=["iso_code", "year"],
        right_on=["territory_iso", "year"],
        how="left"
    )
    merged["attributed_to_iso"] = merged["attributed_to_iso"].fillna(merged["iso_code"])
    merged["coefficient"]       = merged["coefficient"].fillna(1.0)
    merged["attributed_Mt"]     = merged["value_Mt"] * merged["coefficient"]

    attr = (merged.groupby(["attributed_to_iso", "source"])["attributed_Mt"]
            .sum().reset_index()
            .rename(columns={"attributed_to_iso": "iso_code", "attributed_Mt": "value_Mt"}))

    attr_cg = attr.merge(cg[["iso_code"] + CG_GROUP_COLS], on="iso_code", how="left")

    result = {}
    for src, src_df in attr_cg.groupby("source"):
        world_val = float(src_df["value_Mt"].sum())
        result[(src, "world", "World")] = world_val

        for grp_col in CG_GROUP_COLS:
            flagged = src_df[src_df[grp_col] == 1]["value_Mt"]
            result[(src, "group", grp_col)] = float(flagged.sum())

        a1 = result.get((src, "group", "annex_1"), 0.0)
        result[(src, "group", "non_annex_1")] = world_val - a1

        for _, row in src_df.iterrows():
            result[(src, "country", row["iso_code"])] = float(row["value_Mt"])

    _colonial_cache[key] = result
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# TAB VALUE COMPUTATION
# ═══════════════════════════════════════════════════════════════════════════════

def pct(val, world):
    if np.isnan(val) or np.isnan(world) or world == 0:
        return np.nan
    return val / world


def build_tab(tab_name, s, e, is_colonial, measure):
    sources   = tab_sources(s, measure)
    src_codes = [sc for sc, _ in sources]

    if is_colonial:
        col_data = compute_colonial(measure, s, e)

    def get_world_sum(src):
        if src == "GMST":
            return gmst_delta_group("World", measure, s, e)
        if is_colonial:
            v = col_data.get((src, "world", "World"), np.nan)
            return v if v != 0 else np.nan
        return cum_group("World", src, measure, s, e)

    def get_world_all(src):
        if src == "GMST":
            return gmst_delta_group("World", measure, s, e)  # no transport in GMST
        # Colonial tabs: same transport logic as regular tabs.
        return world_all_Mt(src, measure, s, e)

    def get_group(grp_code, src, world_sum):
        if src == "GMST":
            v = gmst_delta_group(grp_code, measure, s, e)
        elif is_colonial:
            v = col_data.get((src, "group", grp_code), np.nan)
        else:
            v = cum_group(grp_code, src, measure, s, e)
        return pct(v, world_sum)

    def get_country(iso, src, world_sum):
        if src == "GMST":
            v = gmst_delta_country(iso, measure, s, e)
        elif is_colonial:
            v = col_data.get((src, "country", iso), np.nan)
        else:
            v = cum_country(iso, src, measure, s, e)
        return pct(v, world_sum)

    def get_nai_g20(src, world_sum):
        if src == "GMST":
            v = gmst_delta_nai_g20(measure, s, e)
        elif is_colonial:
            vals = [col_data.get((src, "country", iso), np.nan) for iso in NAI_G20_ISOS]
            vals = [v for v in vals if not np.isnan(v)]
            v = float(sum(vals)) if vals else np.nan
        else:
            v = cum_nai_g20(src, measure, s, e)
        return pct(v, world_sum)

    world_sums = {src: get_world_sum(src) for src in src_codes}

    row_values = []
    for row_type, code, label in ROW_DEFS:
        if row_type in ("section", "spacer"):
            row_values.append((row_type, code, label, {}))
            continue

        vals = {}
        for src in src_codes:
            ws_val = world_sums.get(src, np.nan)

            if row_type == "world_all":
                vals[src] = get_world_all(src)
            elif row_type == "world_sum":
                vals[src] = world_sums[src]
            elif row_type == "group":
                vals[src] = get_group(code, src, ws_val)
            elif row_type == "country":
                vals[src] = get_country(code, src, ws_val)
            elif row_type == "custom" and code == "NAI_G20":
                vals[src] = get_nai_g20(src, ws_val)

        row_values.append((row_type, code, label, vals))

    return sources, row_values


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL WRITING — DATA TABS
# ═══════════════════════════════════════════════════════════════════════════════

def write_data_tab(wb, tab_name, s, e, is_colonial, measure, sources, row_values):
    ck = colour_key(s, is_colonial)
    hdr_fill, hdr_text, sec_fill, sec_text, tab_col = COLOUR_SCHEME[ck]

    ws = wb.create_sheet(tab_name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_col

    n_src = len(sources)
    ws.column_dimensions["A"].width = 36
    for j in range(2, n_src + 2):
        ws.column_dimensions[get_column_letter(j)].width = 16

    # ── Row 1: title ─────────────────────────────────────────────────────────
    extra = " (colonial attribution)" if is_colonial else ""
    title = f"{tab_name.split(' - ')[0]}  |  Cumulative {s}–{e}{extra}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_src + 1)
    c = ws.cell(row=1, column=1, value=title)
    c.font      = _f(bold=True, size=12, color=hdr_text, font_name=FONT_HEAD)
    c.fill      = _fill(hdr_fill)
    c.alignment = ALIGN_L
    ws.row_dimensions[1].height = 20

    # ── Row 2: column headers ─────────────────────────────────────────────────
    ws.cell(row=2, column=1, value="Group / Country").font      = _f(bold=True, color=hdr_text, font_name=FONT_HEAD)
    ws.cell(row=2, column=1).fill      = _fill(hdr_fill)
    ws.cell(row=2, column=1).alignment = ALIGN_L
    for j, (_, src_label) in enumerate(sources, start=2):
        c = ws.cell(row=2, column=j, value=src_label)
        c.font      = _f(bold=True, color=hdr_text, font_name=FONT_HEAD)
        c.fill      = _fill(hdr_fill)
        c.alignment = ALIGN_C
    ws.freeze_panes = "A3"

    # ── Data rows ─────────────────────────────────────────────────────────────
    current_row = 3
    for row_type, code, label, vals in row_values:

        if row_type == "spacer":
            ws.row_dimensions[current_row].height = 6
            current_row += 1
            continue

        if row_type == "section":
            for j in range(1, n_src + 2):
                ws.cell(row=current_row, column=j).fill = _fill(sec_fill)
            c = ws.cell(row=current_row, column=1, value=label)
            c.font      = _f(bold=True, color=sec_text, font_name=FONT_HEAD)
            c.alignment = ALIGN_L
            current_row += 1
            continue

        is_world = row_type in ("world_all", "world_sum")

        c = ws.cell(row=current_row, column=1)
        c.alignment = ALIGN_L
        if is_world:
            for j in range(1, n_src + 2):
                ws.cell(row=current_row, column=j).fill = _fill(GREY_LIGHT)
            c.value = f"{label}  (GtCO2e)"
            c.font  = _f(bold=True, font_name=FONT_BODY)
        else:
            c.value = label
            c.font  = _f(font_name=FONT_BODY)

        for j, (src_code, _) in enumerate(sources, start=2):
            v    = vals.get(src_code, np.nan)
            cell = ws.cell(row=current_row, column=j)
            cell.alignment = ALIGN_R
            if v is None or (isinstance(v, float) and np.isnan(v)):
                cell.value = EM_DASH
                cell.font  = _f(color=TEXT_MUTED, font_name=FONT_BODY)
            elif is_world:
                if src_code == "GMST":
                    cell.value         = v          # already °C
                    cell.number_format = "0.000"
                else:
                    cell.value         = v / 1000   # Mt → Gt
                    cell.number_format = "0.00"
                cell.font = _f(bold=True, font_name=FONT_BODY)
            else:
                cell.value         = v
                cell.number_format = "0.0%"
                cell.font          = _f(font_name=FONT_BODY)

        current_row += 1


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL WRITING — README TAB
# ═══════════════════════════════════════════════════════════════════════════════

def write_readme_tab(wb, all_tabs):
    ws = wb.create_sheet("README", 0)   # insert at position 0 (first)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = CGD_DARK_TEAL

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 50

    r = 1
    # ── Title ─────────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1,
                value="GHG Emissions Comparative Summary Tables")
    c.font      = _f(bold=True, size=14, color=WHITE, font_name=FONT_HEAD)
    c.fill      = _fill(CGD_DARK_TEAL)
    c.alignment = ALIGN_L
    ws.row_dimensions[r].height = 24
    r += 1

    # Subtitle
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1,
                value="Cumulative emissions shares — 12 static summary tabs")
    c.font      = _f(italic=True, size=10, color=WHITE, font_name=FONT_BODY)
    c.fill      = _fill(CGD_DARK_TEAL)
    c.alignment = ALIGN_L
    ws.row_dimensions[r].height = 16
    r += 2   # blank row

    # ── About section ─────────────────────────────────────────────────────────
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1, value="About")
    c.font      = _f(bold=True, color=WHITE, font_name=FONT_HEAD)
    c.fill      = _fill(CGD_DARK_TEAL)
    c.alignment = ALIGN_L
    r += 1

    about_text = (
        "Each tab shows cumulative emissions for a fixed measure and period. "
        "The two World rows (GtCO2e) give the global total with and without "
        "international transport/bunker fuels. All other rows express the "
        "group or country’s share as a percentage of World (country-attributable). "
        "Colonial tabs reassign pre-independence territorial emissions to "
        "the attributing colonial power using the Carbon Brief territorial-rule database."
    )
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1, value=about_text)
    c.font      = _f(size=10, font_name=FONT_BODY)
    c.alignment = ALIGN_LW
    ws.row_dimensions[r].height = 48
    r += 2

    # ── Contents table ─────────────────────────────────────────────────────────
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1, value="Contents")
    c.font      = _f(bold=True, color=WHITE, font_name=FONT_HEAD)
    c.fill      = _fill(CGD_DARK_TEAL)
    c.alignment = ALIGN_L
    r += 1

    # Column headers for contents table
    hdr_data = ["Tab", "Period", "Type", "Measure", "Sources"]
    for col_idx, hdr in enumerate(hdr_data, start=1):
        c = ws.cell(row=r, column=col_idx, value=hdr)
        c.font      = _f(bold=True, color=WHITE, font_name=FONT_HEAD, size=9)
        c.fill      = _fill(CGD_MED_TEAL)
        c.alignment = ALIGN_C
    ws.row_dimensions[r].height = 14
    r += 1

    # Tab rows with hyperlinks
    for tab_name, s, e, is_colonial, measure in all_tabs:
        ck        = colour_key(s, is_colonial)
        tab_color = COLOUR_SCHEME[ck][4]
        period    = f"{s}–{e}"
        tab_type  = "Colonial" if is_colonial else "As-is"
        srcs      = ", ".join(lbl for _, lbl in tab_sources(s, measure))
        measure_label = next(m for m, lbl in MEASURES if m == measure)
        measure_label = next(lbl for m, lbl in MEASURES if m == measure)

        row_data = [tab_name, period, tab_type, measure_label, srcs]
        for col_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font      = _f(size=9, font_name=FONT_BODY)
            c.alignment = ALIGN_L if col_idx in (1, 4, 5) else ALIGN_C
            if col_idx == 1:
                # Hyperlink to the tab — quote tab name in case it has special chars
                safe_name = tab_name.replace("'", "''")
                c.hyperlink = f"#'{safe_name}'!A1"
                c.font      = Font(color="0563C1", underline="single",
                                   name=FONT_BODY, size=9)
        # Alternate row shading
        if r % 2 == 0:
            for col_idx in range(1, 6):
                ws.cell(row=r, column=col_idx).fill = _fill(CGD_NW_TEAL)
        ws.row_dimensions[r].height = 14
        r += 1

    r += 1  # blank row

    # ── Notes section ─────────────────────────────────────────────────────────
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1, value="Notes")
    c.font      = _f(bold=True, color=WHITE, font_name=FONT_HEAD)
    c.fill      = _fill(CGD_DARK_TEAL)
    c.alignment = ALIGN_L
    r += 1

    notes = [
        ("Climate Watch (1990–2024 tabs)",
         "Climate Watch data runs through 2023 only. The 2024 year is absent from all "
         "Climate Watch figures in the 1990–2024 tabs. For this analysis, "
         "Climate Watch is treated as complete to 2024; the missing year "
         "causes a slight undercount relative to other sources."),
        ("Climate Watch vs. other sources",
         "Climate Watch uses UNFCCC national GHG inventory data, while OWID uses the "
         "Global Carbon Budget (fossil CO2) methodology. The direction of divergence "
         "depends on the measure: for GHG excl LULUCF (1990–2024) CW is ~10% higher "
         "than OWID but ~5% lower than PRIMAP/EDGAR; for GHG incl LULUCF CW is lower "
         "than OWID. Separate from the missing-2024 effect."),
        ("World (all emissions) sources",
         "Transport component by source: OWID — World entity in raw CSV (includes "
         "international aviation + shipping). EDGAR — AIR + SEA entities added to "
         "country sum. Climate Watch — WORLD entity (includes bunker fuels). "
         "GCP fossil — ‘World’ column in raw xlsx (includes shipping + aviation). "
         "PRIMAP: no international transport in dataset; "
         "World (all) = World (country-attributable) for this source. "
         "GCP LUC variants (BLUE/OSCAR/LUCE): GCP fossil transport component added "
         "to the country-sum (LUC models have no transport; delta from GCP fossil "
         "'World' column vs country sum, ~46 Gt for 1850–2024). "
         "OWID co2_incLUC: World entity LUC differs from country-sum LUC — "
         "fossil transport delta added to country sum instead of using World entity directly."),
        ("Colonial attribution (tabs 9–12)",
         "Emissions are reassigned from territories to attributing powers using "
         "the Carbon Brief territorial-rule database (1850–2023, forward-filled to 2024). "
         "Each territory-year’s emissions are split across powers by the coefficient "
         "in that database (coefficients sum to ~1.0 per territory-year). "
         "Groups show totals attributed TO countries in each group. "
         "Countries not covered by the database are attributed to themselves (coefficient=1). "
         "World (all emissions) for colonial tabs is the same as the as-is 1850–2024 tabs "
         "(transport is a global total; colonial attribution does not reallocate it)."),
        ("Non-Annex I G20 members",
         "Non-Annex I G20 members row = sum of the nine G20 member states outside "
         "Annex I: Argentina (ARG), Brazil (BRA), China (CHN), India (IND), "
         "Indonesia (IDN), Mexico (MEX), Saudi Arabia (SAU), South Africa (ZAF), "
         "and South Korea (KOR). In colonial tabs, emissions from these territories "
         "prior to independence are attributed to the colonial power under the "
         "Carbon Brief database; post-independence emissions to the states themselves."),
        ("GCP LUC variants",
         "GCP BLUE, OSCAR, and LUCE are alternative land-use-change model estimates. "
         "Each is combined with GCP fossil emissions to produce CO2 incl LULUCF. "
         "GCP sources cover CO2 only (no all-GHG equivalent). "
         "OWID's country-level land_use_change_co2 values match GCP BLUE exactly — "
         "OWID uses BLUE for national attribution. "
         "However, OWID's World entity row uses the GCB global budget LUC, which is a "
         "consensus estimate across multiple models and is not equal to the country sum of BLUE. "
         "As a result, OWID World(all) co2_incLUC uses country-sum LUC (BLUE-based) + "
         "fossil transport delta — not the World entity's LUC directly."),
        ("GMST column",
         "GMST shows the cumulative temperature response (Jones et al. 2025) attributable "
         "to each group or country. World rows show total GMST in °C for the tab's period; "
         "all other rows show GMST share as % of the World total. "
         "For 1990–2024 tabs, GMST is computed as the delta from 1989 to 2024. "
         "Colonial tabs show as-is GMST (not colonial-attributed) — colonial GMST "
         "attribution requires a differencing approach not applied here."),
    ]

    for heading, body in notes:
        c = ws.cell(row=r, column=1, value=heading)
        c.font      = _f(bold=True, size=9, color=TEXT_DARK, font_name=FONT_BODY)
        c.alignment = ALIGN_L
        c.fill      = _fill(CGD_NW_TEAL)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        r += 1

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(row=r, column=1, value=body)
        c.font      = _f(size=9, font_name=FONT_BODY)
        c.alignment = ALIGN_LW
        ws.row_dimensions[r].height = 42
        r += 1

    # ── Source coverage table ─────────────────────────────────────────────────
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws.cell(row=r, column=1, value="Source coverage")
    c.font      = _f(bold=True, color=WHITE, font_name=FONT_HEAD)
    c.fill      = _fill(CGD_DARK_TEAL)
    c.alignment = ALIGN_L
    r += 1

    src_hdrs = ["Source", "Year range", "Measures",
                "World (all) = World (ctry-attr.)?", "Transport component"]
    for col_idx, hdr in enumerate(src_hdrs, start=1):
        c = ws.cell(row=r, column=col_idx, value=hdr)
        c.font      = _f(bold=True, color=WHITE, font_name=FONT_HEAD, size=9)
        c.fill      = _fill(CGD_MED_TEAL)
        c.alignment = ALIGN_C
    ws.row_dimensions[r].height = 28
    r += 1

    for _, src_label, yr_range, measures, transport in SOURCE_META:
        same = "Yes" if "None" in transport else "No"
        row_data = [src_label, yr_range, measures, same, transport]
        for col_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font      = _f(size=9, font_name=FONT_BODY)
            c.alignment = ALIGN_LW if col_idx in (2, 3, 5) else ALIGN_C
        if r % 2 == 0:
            for col_idx in range(1, 6):
                ws.cell(row=r, column=col_idx).fill = _fill(CGD_NW_TEAL)
        ws.row_dimensions[r].height = 28
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

print("\nComputing tab values ...")
tab_data = []
for tab_name, s, e, is_colonial, measure in TABS:
    print(f"  {tab_name} ...")
    sources, row_values = build_tab(tab_name, s, e, is_colonial, measure)
    tab_data.append((tab_name, s, e, is_colonial, measure, sources, row_values))

print("\nBuilding workbook ...")
wb = openpyxl.Workbook()
wb.remove(wb.active)

# README first
write_readme_tab(wb, TABS)
print("  Written: README")

# Data tabs
for tab_name, s, e, is_colonial, measure, sources, row_values in tab_data:
    write_data_tab(wb, tab_name, s, e, is_colonial, measure, sources, row_values)
    print(f"  Written: {tab_name}")

os.makedirs(OUT_DIR, exist_ok=True)
wb.save(OUT_XLSX)
print(f"\nSaved: {os.path.basename(OUT_XLSX)}  (README + {len(TABS)} data tabs)")

import shutil
shutil.copyfile(OUT_XLSX, ANNEX_C_XLSX)
print(f"Copied to paper deliverable: {os.path.basename(ANNEX_C_XLSX)}")
