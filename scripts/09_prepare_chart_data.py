"""
09_prepare_chart_data.py
Prepare data for all chart figures in the GHG comparative analysis paper.

Inputs (from data/outputs/):
  groups_summary.csv
  countries_annual.csv

Inputs (raw):
  CW_HistoricalEmissions_ClimateWatch.csv  (for Fig 1 gas/sector breakdown)

Outputs (to data/outputs/charts/):
  fig1_data.csv  — CW 2023 WORLD: gas breakdown + sector breakdown (% shares)
  fig2_data.csv  — OWID World: CO2/non-CO2 x fossil/LULUCF, 4 series, 1850–2024
  fig3_data.csv  — Cumulative shares 1850–2024 (OWID, PRIMAP, GCP; no EDGAR/CW)
  fig4_data.csv  — Cumulative shares 1990–2024 (all sources incl EDGAR, CW)
  fig5_data.csv  — OWID annual by 7 derived country groups, 1850–2024
  fig6_data.csv  — Cumulative share by end year (forward cumsum from 1850)
  fig7_data.csv  — Cumulative share by start year (reverse cumsum to 2024)
  fig8_data.csv  — Colonial-attributed cumulative shares 1850–2024
  fig9_data.csv  — Annual per-capita emissions, 4 groups x 4 measures
  fig10_data.csv / fig11_data.csv — Cumulative per-capita (1850- / 1990-baseline)
  interactive_[measure].csv (+ _colonial) and interactive_gmst.csv (+ _colonial)
    — wide-format inputs to script 12 (interactive payload builder)

Dumbbell row structure (figs 3, 4):
  8 data rows = 4 measures x 2 groups (Annex II, Annex I), with spacers between.
  Columns: y_pos, label, owid, primap, gcp_fossil, gcp_blue, gcp_oscar, gcp_luce,
           edgar, cw, gmst_ref
  NaN where a source does not cover a given measure or period.

MultiLine structure (figs 6, 7):
  Group A = annex_2 (Annex II), Group B = annex_1 (Annex I).
  Columns: measure, year,
           annex2_[src], annex1_[src], non_annex1_[src]  — annual MtCO2e per active source
           annex2_GMST, annex1_GMST                       — cumulative GMST share % (Python-computed)
  Cumulative sums and % shares are computed by Excel formulas in the vals tab.
  World = annex_1 + non_annex_1 is derived in Excel (not in the CSV).
  Source slots vary by measure (see SOURCES_D / SOURCES_E dicts).
"""

import os
import numpy as np
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW_DIR     = os.path.join(QA_DIR, "data", "raw")
OUT_DIR     = os.path.join(QA_DIR, "data", "outputs")
CHART_DIR   = os.path.join(OUT_DIR, "charts")
os.makedirs(CHART_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# IRL pre-1950 population patch helpers (T2: Fig G / Fig H)
# ---------------------------------------------------------------------------
def _load_irl_population_patch():
    """Return IRL population patch DataFrame for 1850-1949 (years NaN in owid-co2-data.csv).

    1850-1920: OWID population.csv "Ireland (whole island)" — whole island, no ISO code.
               Minor overcount (~1.3M NI of ~6M total) but far preferable to excluding IRL.
    1921-1949: Maddison Project Database 2023, Full data sheet, countrycode=IRL
               (Republic of Ireland only from 1921; units: thousands → ×1000).
    GBR population confirmed independent — no double-counting risk.
    See docs/methodology-notes.md §"Fig G per-capita: NaN handling".
    """
    pop_csv = pd.read_csv(os.path.join(RAW_DIR, "population.csv"), low_memory=False)
    wi = pop_csv[pop_csv["Entity"] == "Ireland (whole island)"][["Year", "Population"]].copy()
    wi = wi.dropna(subset=["Population"])
    wi = wi[wi["Year"].between(1850, 1920)].rename(
        columns={"Year": "year", "Population": "population"})
    wi["iso_code"] = "IRL"

    mad = pd.read_excel(os.path.join(RAW_DIR, "mpd2023_web.xlsx"),
                        sheet_name="Full data", usecols=["countrycode", "year", "pop"])
    mad = mad[(mad["countrycode"] == "IRL") & (mad["year"].between(1921, 1949))].copy()
    mad["population"] = mad["pop"] * 1000
    mad["iso_code"] = "IRL"
    mad = mad[["iso_code", "year", "population"]]

    patch = pd.concat([wi[["iso_code", "year", "population"]], mad], ignore_index=True)
    patch["year"] = patch["year"].astype(int)
    return patch


def _apply_irl_patch(pop_df, patch):
    """Coalesce IRL pre-1950 population patch into a population DataFrame."""
    merged = pop_df.merge(
        patch.rename(columns={"population": "_irl_pop"}),
        on=["iso_code", "year"], how="left"
    )
    merged["population"] = merged["population"].combine_first(merged["_irl_pop"])
    return merged.drop(columns=["_irl_pop"])


# ---------------------------------------------------------------------------
# Load pipeline outputs
# ---------------------------------------------------------------------------
print("Loading groups_summary.csv ...")
grp = pd.read_csv(os.path.join(OUT_DIR, "groups_summary.csv"), low_memory=False)
grp["year"] = grp["year"].astype(int)

print("Loading countries_annual.csv ...")
cntry = pd.read_csv(os.path.join(OUT_DIR, "countries_annual.csv"), low_memory=False)
cntry["year"] = cntry["year"].astype(int)

print("Loading CW raw data ...")
cw_raw = pd.read_csv(os.path.join(RAW_DIR, "CW_HistoricalEmissions_ClimateWatch.csv"),
                     low_memory=False)
cw_raw["Sector"] = cw_raw["Sector"].str.strip()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
MEASURES = ["co2_excLUC", "co2_incLUC", "ghg_excLUC", "ghg_incLUC"]

MEASURE_LABELS = {
    "co2_excLUC": "CO2 excl. LULUCF",
    "co2_incLUC": "CO2 incl. LULUCF",
    "ghg_excLUC": "All GHG excl. LULUCF",
    "ghg_incLUC": "All GHG incl. LULUCF",
}

# Sources included per measure for Fig A (1850–2024; excludes EDGAR 1970+, CW 1990+)
SOURCES_A = {
    "co2_excLUC": ["OWID", "PRIMAP", "GCP_fossil"],
    "co2_incLUC": ["OWID", "PRIMAP", "GCP_BLUE", "GCP_OSCAR", "GCP_LUCE"],
    "ghg_excLUC": ["OWID", "PRIMAP"],
    "ghg_incLUC": ["OWID", "PRIMAP"],
}

# Sources per measure for Fig B (1990–2024; includes EDGAR, CW where available)
SOURCES_B = {
    "co2_excLUC": ["OWID", "PRIMAP", "GCP_fossil", "EDGAR"],
    "co2_incLUC": ["OWID", "PRIMAP", "GCP_BLUE", "GCP_OSCAR", "GCP_LUCE", "EDGAR"],
    "ghg_excLUC": ["OWID", "PRIMAP", "EDGAR", "ClimateWatch"],
    "ghg_incLUC": ["OWID", "PRIMAP", "EDGAR"],
}

# Sources for figs D and E (up to 4 per measure, matching multiline template 4-col slots)
SOURCES_D = {
    # FigD: cumulative forward shares from 1850 to each end year.
    # EDGAR (1970+) and CW (1990+) are excluded from ghg_excLUC: in a forward cumsum
    # from 1850, pre-coverage years sum to 0 → the dotted line plots as 0% before
    # 1970/1990, which is misleading. These sources belong in FigE instead.
    "co2_excLUC": ["OWID", "PRIMAP", "GCP_fossil", None],
    "co2_incLUC": ["OWID", "PRIMAP", "GCP_BLUE", "GCP_OSCAR"],
    "ghg_excLUC": ["OWID", "PRIMAP", None, None],
    "ghg_incLUC": ["OWID", "PRIMAP", None, None],
}

SOURCES_E = {
    # FigE: cumulative shares from each start year to 2024. EDGAR (1970+) and
    # ClimateWatch (1990+) are appropriate here: for start years before their coverage
    # the cumsum naturally covers only the available period (a flat line at the
    # 1970–2024 / 1990–2022 total), which is visible and interpretable in the chart.
    "co2_excLUC": ["OWID", "PRIMAP", "GCP_fossil", None],
    "co2_incLUC": ["OWID", "PRIMAP", "GCP_BLUE", "GCP_OSCAR"],
    "ghg_excLUC": ["OWID", "PRIMAP", "EDGAR", "ClimateWatch"],
    "ghg_incLUC": ["OWID", "PRIMAP", "EDGAR", None],
}

# All source keys used in dumbbell columns
ALL_SRC_COLS = ["owid", "primap", "gcp_fossil", "gcp_blue", "gcp_oscar", "gcp_luce",
                "edgar", "cw"]

# Source order used in the wide-format interactive_[measure].csv files
# (and the colonial-attributed _J variants) — the column order in those CSVs
# must match between as-reported and J so the dynamic chart can swap between them.
DB_ALL_SOURCES_RAW = ["OWID", "PRIMAP", "GCP_fossil", "GCP_BLUE", "GCP_OSCAR",
                       "GCP_LUCE", "EDGAR", "ClimateWatch"]

# Map from source key in CSV to column name (lowercase, no hyphens)
SRC_TO_COL = {
    "OWID":         "owid",
    "PRIMAP":       "primap",
    "GCP_fossil":   "gcp_fossil",
    "GCP_BLUE":     "gcp_blue",
    "GCP_OSCAR":    "gcp_oscar",
    "GCP_LUCE":     "gcp_luce",
    "EDGAR":        "edgar",
    "ClimateWatch": "cw",
}

# ---------------------------------------------------------------------------
# Helper: look up a single value from groups_summary
# ---------------------------------------------------------------------------
def grp_val(source, group, measure, year):
    row = grp[
        (grp["source"] == source) & (grp["group"] == group) &
        (grp["measure"] == measure) & (grp["year"] == year)
    ]
    return float(row["value_Mt"].iloc[0]) if not row.empty else np.nan


# ---------------------------------------------------------------------------
# Helper: cumulative share for a period
# ---------------------------------------------------------------------------
def cum_share(source, group, measure, start_yr, end_yr):
    """Return group's share (%) of World cumulative emissions for the period."""
    mask_g = (
        (grp["source"] == source) & (grp["group"] == group) &
        (grp["measure"] == measure) &
        (grp["year"] >= start_yr) & (grp["year"] <= end_yr)
    )
    mask_w = (
        (grp["source"] == source) & (grp["group"] == "World") &
        (grp["measure"] == measure) &
        (grp["year"] >= start_yr) & (grp["year"] <= end_yr)
    )
    g_sum = grp.loc[mask_g, "value_Mt"].sum()
    w_sum = grp.loc[mask_w, "value_Mt"].sum()
    if w_sum == 0 or np.isnan(w_sum):
        return np.nan
    return g_sum / w_sum * 100


# ---------------------------------------------------------------------------
# Helper: GMST warming share for a period
# ---------------------------------------------------------------------------
_GMST_MSR_MAP = {
    "co2_excLUC": "gmst_co2_excLUC",
    "co2_incLUC": "gmst_co2_incLUC",
    "ghg_excLUC": "gmst_ghg_excLUC",
    "ghg_incLUC": "gmst_ghg_incLUC",
}

def gmst_share(group, msr, start_yr, end_yr):
    """Return group's share (%) of World GMST response for the period.

    Uses the per-measure GMST label so CO2 vs GHG and excl/incl LUC each
    use the correct gas/component from the GMST response dataset.
    """
    gmst_msr = _GMST_MSR_MAP[msr]
    g_end = grp_val("GMST", group, gmst_msr, end_yr)
    w_end = grp_val("GMST", "World", gmst_msr, end_yr)
    g_pre = grp_val("GMST", group, gmst_msr, start_yr - 1) if start_yr > 1851 else 0.0
    w_pre = grp_val("GMST", "World", gmst_msr, start_yr - 1) if start_yr > 1851 else 0.0
    denom = w_end - w_pre
    if pd.isna(denom) or denom == 0:
        return np.nan
    numer = g_end - g_pre
    if pd.isna(numer):
        return np.nan
    return numer / denom * 100


# ===========================================================================
# Fig 1: PiePair — CW 2023 World gas + sector breakdown
# ===========================================================================
print("\nBuilding fig1_data.csv ...")

world_2023 = cw_raw[cw_raw["Country"] == "WORLD"][["Gas", "Sector", "2023"]].copy()
world_2023["2023"] = pd.to_numeric(world_2023["2023"], errors="coerce")

total_excl = float(
    world_2023[(world_2023["Gas"] == "All GHG") &
               (world_2023["Sector"] == "Total excluding LULUCF")]["2023"].iloc[0]
)
total_incl = float(
    world_2023[(world_2023["Gas"] == "All GHG") &
               (world_2023["Sector"] == "Total including LULUCF")]["2023"].iloc[0]
)

# Pie 1: by gas (% of total excl. LUCF)
gas_df = world_2023[
    (world_2023["Sector"] == "Total excluding LULUCF") &
    (world_2023["Gas"].isin(["CO2", "CH4", "N2O", "F-Gas"]))
].copy()
gas_df["share_pct"] = gas_df["2023"] / total_excl * 100
gas_df = gas_df.rename(columns={"Gas": "category", "2023": "value_Mt"})
gas_df["breakdown"] = "gas"
gas_df = gas_df[["breakdown", "category", "value_Mt", "share_pct"]]

# Pie 2: by main sector (% of total incl. LUCF)
# Bunker Fuels is a memo item already included within the Energy sector total;
# including it separately would overcount. Use 5 clean sectors only.
main_sectors = [
    "Energy", "Industrial Processes", "Agriculture", "Waste",
    "Land Use, Land-Use Change and Forestry",
]
sec_df = world_2023[
    (world_2023["Gas"] == "All GHG") &
    (world_2023["Sector"].isin(main_sectors))
].copy()
sec_df["share_pct"] = sec_df["2023"] / total_incl * 100
sec_df = sec_df.rename(columns={"Sector": "category", "2023": "value_Mt"})
sec_df["category"] = sec_df["category"].replace(
    "Land Use, Land-Use Change and Forestry", "LULUCF"
)
sec_df["breakdown"] = "sector"
sec_df = sec_df[["breakdown", "category", "value_Mt", "share_pct"]]

fig1 = pd.concat([gas_df, sec_df], ignore_index=True)
fig1.to_csv(os.path.join(CHART_DIR, "fig1_data.csv"), index=False, float_format="%.3f")
print(f"  Written: fig1_data.csv  ({len(fig1)} rows)")


# ===========================================================================
# Fig 2 combined: StackedArea — OWID World, 1850–2024, 4 series in one chart
#   Series: CO2 fossil, non-CO2 fossil, CO2 LULUCF, non-CO2 LULUCF
#   Total stacks to GHG incl. LUC (World, OWID)
# ===========================================================================
print("\nBuilding fig2_data.csv ...")

_src = {}
for _m in ["co2_excLUC", "co2_incLUC", "ghg_excLUC", "ghg_incLUC"]:
    _src[_m] = (
        grp[(grp["source"] == "OWID") & (grp["group"] == "World") &
            (grp["measure"] == _m) & (grp["year"] >= 1850)]
        .set_index("year")["value_Mt"]
    )

_yrs2c = sorted(set(_src["co2_excLUC"].index) | set(_src["ghg_excLUC"].index))

def _gv(msr, y):
    v = _src[msr].get(y, np.nan)
    return float(v) if not pd.isna(v) else np.nan

def _gd(msr_a, msr_b, y):
    a, b = _gv(msr_a, y), _gv(msr_b, y)
    return a - b if not (pd.isna(a) or pd.isna(b)) else np.nan

rows_2c = []
for y in _yrs2c:
    co2_lulucf    = _gd("co2_incLUC", "co2_excLUC", y)
    ghg_lulucf    = _gd("ghg_incLUC", "ghg_excLUC", y)
    nonco2_lulucf = (
        ghg_lulucf - co2_lulucf
        if not (pd.isna(ghg_lulucf) or pd.isna(co2_lulucf)) else np.nan
    )
    rows_2c.append({
        "year":              y,
        "co2_fossil_Mt":     _gv("co2_excLUC", y),
        "nonco2_fossil_Mt":  _gd("ghg_excLUC", "co2_excLUC", y),
        "co2_lulucf_Mt":     co2_lulucf,
        "nonco2_lulucf_Mt":  nonco2_lulucf,
    })

fig2c = pd.DataFrame(rows_2c)
fig2c.to_csv(os.path.join(CHART_DIR, "fig2_data.csv"), index=False, float_format="%.3f")
print(f"  Written: fig2_data.csv  ({len(fig2c)} rows)")


# ===========================================================================
# Figs A & B: Dumbbell
# ===========================================================================
def build_dumbbell(start_yr, end_yr, src_map):
    """Build dumbbell DataFrame for a given period and source map.

    Returns DataFrame with columns:
      y_pos, label, owid, primap, gcp_fossil, gcp_blue, gcp_oscar, gcp_luce,
      edgar, cw, gmst_ref

    Row order (top-to-bottom in chart, i.e. highest y_pos first visually):
      GHG incl LUC, CO2 incl LUC, GHG excl LUC, CO2 excl LUC.
      Within each pair: Annex I first (lower y_pos), Annex II second (higher
      y_pos) so Annex II appears above in the Excel chart.
    """
    rows = []
    y = 1
    # Annex I first (lower y_pos = bottom of chart), then Annex II (higher =
    # visually above) so each pair reads Annex II above Annex I.
    annex_groups = [
        ("annex_1", "Annex I"),
        ("annex_2", "Annex II"),
    ]

    # Dumbbell measure order per JB: top-to-bottom = GHG incl, CO2 incl,
    # GHG excl, CO2 excl.  In the CSV the first row has the lowest y_pos
    # (bottom of chart), so we reverse: CO2 excl first → GHG incl last.
    dumbbell_measures = [
        ("co2_excLUC", MEASURE_LABELS["co2_excLUC"]),
        ("ghg_excLUC", MEASURE_LABELS["ghg_excLUC"]),
        ("co2_incLUC", MEASURE_LABELS["co2_incLUC"]),
        ("ghg_incLUC", MEASURE_LABELS["ghg_incLUC"]),
    ]
    for idx, (msr, msr_label) in enumerate(dumbbell_measures):
        for grp_key, grp_label in annex_groups:
            row = {"y_pos": y, "label": f"{msr_label} — {grp_label}"}
            active_srcs = src_map.get(msr, [])
            for src_name, col_name in SRC_TO_COL.items():
                if src_name in active_srcs:
                    row[col_name] = cum_share(src_name, grp_key, msr, start_yr, end_yr)
                else:
                    row[col_name] = np.nan
            row["gmst_ref"] = gmst_share(grp_key, msr, start_yr, end_yr)
            rows.append(row)
            y += 1
        # Spacer between measure groups (not after the last one)
        if idx < len(dumbbell_measures) - 1:
            rows.append({"y_pos": y, "label": ""})
            y += 1

    cols = ["y_pos", "label"] + ALL_SRC_COLS + ["gmst_ref"]
    df = pd.DataFrame(rows, columns=cols)
    return df


print("\nBuilding fig3_data.csv (1850–2024) ...")
figA = build_dumbbell(1850, 2024, SOURCES_A)
figA.to_csv(os.path.join(CHART_DIR, "fig3_data.csv"), index=False, float_format="%.2f")
print(f"  Written: fig3_data.csv  ({len(figA)} rows)")

print("\nBuilding fig4_data.csv (1990–2024) ...")
figB = build_dumbbell(1990, 2024, SOURCES_B)
figB.to_csv(os.path.join(CHART_DIR, "fig4_data.csv"), index=False, float_format="%.2f")
print(f"  Written: fig4_data.csv  ({len(figB)} rows)")


# ===========================================================================
# Fig C: StackedArea — OWID annual by 7 derived country groups, 1850–2024
#
# Stack order (bottom to top): USA, EU, Other_AII, Other_AI, CHN, IND, Other_nonAI
#   USA       = country USA (from countries_annual)
#   EU        = group eu (from groups_summary)
#   Other_AII = annex_2 - USA - EU
#   Other_AI  = annex_1 - annex_2
#   CHN       = country CHN
#   IND       = country IND
#   Other_nonAI = World - annex_1 - CHN - IND
# ===========================================================================
print("\nBuilding fig5_data.csv ...")

# European Annex 2 = all Annex 2 except USA, AUS, CAN, JPN, NZL (18 European members).
# Other Annex 2   = AUS, CAN, JPN, NZL.
# Other non-Annex I G20 = non-Annex I G20 members excl. CHN and IND, which have
# their own series (split from Other non-Annex I). Replaced Gulf states 2026-07-07.
_OTHER_A2_ISOS  = ["AUS", "CAN", "JPN", "NZL"]
_NAI_G20_ISOS   = ["ARG", "BRA", "IDN", "KOR", "MEX", "SAU", "ZAF"]

chunks_c = []
for msr in MEASURES:
    def gs(group):
        """Annual OWID series for a group, indexed by year."""
        return (
            grp[(grp["source"] == "OWID") & (grp["group"] == group) &
                (grp["measure"] == msr) & (grp["year"] >= 1850)]
            .set_index("year")["value_Mt"]
        )

    def cs(iso):
        """Annual OWID series for a country, indexed by year."""
        return (
            cntry[(cntry["iso_code"] == iso) & (cntry["source"] == "OWID") &
                  (cntry["measure"] == msr) & (cntry["year"] >= 1850)]
            .set_index("year")["value_Mt"]
        )

    def cs_sum(isos):
        """Sum annual OWID series across a list of ISO codes."""
        srs = [cs(iso) for iso in isos]
        combined = pd.concat(srs).groupby(level=0).sum()
        return combined

    s_world    = gs("World")
    s_aii      = gs("annex_2")
    s_ai       = gs("annex_1")
    s_usa      = cs("USA")
    s_chn      = cs("CHN")
    s_ind      = cs("IND")
    s_other_a2 = cs_sum(_OTHER_A2_ISOS)
    s_nai_g20  = cs_sum(_NAI_G20_ISOS)

    years_c = sorted(s_world.index)

    def v(series, y):
        val = series.get(y, np.nan)
        return float(val) if not pd.isna(val) else np.nan

    def safe_sub(total, *parts):
        """Subtract parts from total; treat NaN parts as 0; return NaN if total is NaN."""
        if pd.isna(total):
            return np.nan
        return total - sum(0 if pd.isna(p) else p for p in parts)

    rows_c = []
    for y in years_c:
        usa_v      = v(s_usa,      y)
        aii_v      = v(s_aii,      y)
        ai_v       = v(s_ai,       y)
        chn_v      = v(s_chn,      y)
        ind_v      = v(s_ind,      y)
        world_v    = v(s_world,    y)
        other_a2_v = v(s_other_a2, y)
        nai_g20_v  = v(s_nai_g20,  y)
        # European Annex 2 = Annex 2 minus USA minus Other Annex 2
        euro_a2_v  = safe_sub(aii_v, usa_v, other_a2_v)
        rows_c.append({
            "measure":      msr,
            "year":         y,
            "usa":          usa_v,
            "euro_a2":      euro_a2_v,
            "other_a2":     other_a2_v,
            "other_ai":     safe_sub(ai_v, aii_v),
            "chn":          chn_v,
            "ind":          ind_v,
            "nai_g20":      nai_g20_v,
            "other_non_ai": safe_sub(world_v, ai_v, chn_v, ind_v, nai_g20_v),
        })
    chunks_c.append(pd.DataFrame(rows_c))

figC = pd.concat(chunks_c, ignore_index=True)
figC.to_csv(os.path.join(CHART_DIR, "fig5_data.csv"), index=False, float_format="%.3f")
print(f"  Written: fig5_data.csv  ({len(figC)} rows, {figC['measure'].nunique()} measures)")


# ===========================================================================
# Figs D & E: MultiLine — cumulative shares by end year (D) and start year (E)
#
# Output columns per figure:
#   measure, year,
#   annex2_src1, annex2_src2, annex2_src3, annex2_src4,
#   annex1_src1, annex1_src2, annex1_src3, annex1_src4
#
# Where src1–src4 map to the source list in SOURCES_D / SOURCES_E for that measure.
# None entries produce NaN columns.
# Column names embed the actual source name (e.g. annex2_OWID, annex2_GCP_fossil).
# All source columns are present in each row; NaN where the measure/source doesn't apply.
# ===========================================================================

def build_multiline(sources_map, direction, start_yr=1850):
    """Build multiline DataFrame with annual MtCO2e values per source/group.

    direction = 'forward' -> FigD: GMST cumulative from start_yr to each end year.
    direction = 'reverse' -> FigE: GMST cumulative from each start year to 2024.

    Emission columns are annual MtCO2e absolute values (NOT percentages).
    Cumulative sums and % shares are computed by Excel formulas in the calcs tab.
    GMST columns remain Python-computed cumulative % (GMST is in °C, not
    decomposable into annual values for re-summing in Excel).

    CSV columns per row:
        measure, year
        annex2_[src], annex1_[src], non_annex1_[src]  — MtCO2e per active source
        annex2_GMST, annex1_GMST                       — cumulative GMST share % (Python-computed)

    World denominator is NOT in the CSV. It is derived in Excel as:
        cum_world = cum_annex1 + cum_non_annex1
    This makes the two-step derivation explicit and human-verifiable.
    """
    END_YR = 2024

    # Precompute GMST cumulative series (°C) per measure — each uses its correct gas/component.
    # Measure → GMST measure label mapping (must match script 06 output).
    GMST_MSR_MAP = {
        "co2_excLUC": "gmst_co2_excLUC",
        "co2_incLUC": "gmst_co2_incLUC",
        "ghg_excLUC": "gmst_ghg_excLUC",
        "ghg_incLUC": "gmst_ghg_incLUC",
    }
    gmst_series = {}  # keyed (msr, ann_grp)
    for msr, gmst_msr in GMST_MSR_MAP.items():
        for ann_grp in ["annex_2", "annex_1", "World"]:
            gmst_series[(msr, ann_grp)] = (
                grp[(grp["source"] == "GMST") & (grp["group"] == ann_grp) &
                    (grp["measure"] == gmst_msr)]
                .set_index("year")["value_Mt"]
                .sort_index()
            )

    def _gmst_share_at(msr, ann_grp, yr_start, yr_end):
        """GMST share (%) for period [yr_start, yr_end] using the correct gas/component for msr."""
        key_g = (msr, ann_grp)
        key_w = (msr, "World")
        if key_g not in gmst_series or key_w not in gmst_series:
            return np.nan
        g = gmst_series[key_g]
        w = gmst_series[key_w]
        g_end = g.get(yr_end, np.nan)
        w_end = w.get(yr_end, np.nan)
        g_pre = g.get(yr_start - 1, 0.0) if yr_start > 1851 else 0.0
        w_pre = w.get(yr_start - 1, 0.0) if yr_start > 1851 else 0.0
        denom = w_end - w_pre
        if pd.isna(denom) or denom == 0:
            return np.nan
        numer = g_end - g_pre
        return numer / denom * 100 if not pd.isna(numer) else np.nan

    chunks_ml = []

    for msr in MEASURES:
        srcs = sources_map[msr]  # list of up to 4 sources (may contain None)

        # Precompute annual series for each active source × group pair.
        # non_annex_1 is now in groups_summary; use it directly.
        grp_series = {}
        for src in srcs:
            if src is None:
                continue
            grp_series[src] = {}
            for ann_grp in ["annex_2", "annex_1", "non_annex_1"]:
                s = (
                    grp[(grp["source"] == src) & (grp["group"] == ann_grp) &
                        (grp["measure"] == msr)]
                    .set_index("year")["value_Mt"]
                    .sort_index()
                )
                grp_series[src][ann_grp] = s

        # All years with data in [start_yr, END_YR]
        all_years = sorted(
            y for y in set().union(*[s.index for src_d in grp_series.values()
                                      for s in src_d.values()])
            if start_yr <= y <= END_YR
        ) if grp_series else []

        rows_ml = []
        for y in all_years:
            row = {"measure": msr, "year": y}
            for src in srcs:
                if src is None:
                    continue
                # Annual MtCO2e absolute values; cumulative sums computed in Excel.
                # Column name mapping: annex_2→annex2, annex_1→annex1, non_annex_1→non_annex1
                for ann_grp in ["annex_2", "annex_1", "non_annex_1"]:
                    pfx = ann_grp.replace("_", "")   # annex2, annex1, nonannex1 — wait
                    # Keep underscores for clarity: annex2, annex1, non_annex1
                    pfx = {"annex_2": "annex2", "annex_1": "annex1",
                           "non_annex_1": "non_annex1"}[ann_grp]
                    row[f"{pfx}_{src}"] = grp_series[src][ann_grp].get(y, np.nan)

            # GMST: cumulative % (direction-dependent; per-measure gas/component)
            for ann_grp in ["annex_2", "annex_1"]:
                pfx = {"annex_2": "annex2", "annex_1": "annex1"}[ann_grp]
                if direction == "forward":
                    row[f"{pfx}_GMST"] = _gmst_share_at(msr, ann_grp, start_yr, y)
                else:  # reverse
                    row[f"{pfx}_GMST"] = _gmst_share_at(msr, ann_grp, y, END_YR)

            rows_ml.append(row)

        chunks_ml.append(pd.DataFrame(rows_ml))

    result = pd.concat(chunks_ml, ignore_index=True)
    # Column order: measure, year, per-source triplets (annex2/annex1/non_annex1), then GMST
    base_cols = ["measure", "year"]
    seen_srcs = []
    for msr in MEASURES:
        for src in sources_map[msr]:
            if src is not None and src not in seen_srcs:
                seen_srcs.append(src)
    src_cols = []
    for src in seen_srcs:
        for pfx in ["annex2", "annex1", "non_annex1"]:
            col = f"{pfx}_{src}"
            if col in result.columns:
                src_cols.append(col)
    gmst_cols = [c for c in result.columns if c.endswith("_GMST")]
    result = result.reindex(columns=base_cols + src_cols + gmst_cols)
    return result


print("\nBuilding fig6_data.csv (annual MtCO2e; FigD = forward cumul from 1850) ...")
figD = build_multiline(SOURCES_D, "forward", start_yr=1850)
figD.to_csv(os.path.join(CHART_DIR, "fig6_data.csv"), index=False, float_format="%.3f")
print(f"  Written: fig6_data.csv  ({len(figD)} rows, {figD['measure'].nunique()} measures)")

print("\nBuilding fig7_data.csv (annual MtCO2e; FigE = reverse cumul to 2024) ...")
figE = build_multiline(SOURCES_E, "reverse", start_yr=1850)
figE.to_csv(os.path.join(CHART_DIR, "fig7_data.csv"), index=False, float_format="%.3f")
print(f"  Written: fig7_data.csv  ({len(figE)} rows, {figE['measure'].nunique()} measures)")


# ===========================================================================
# Fig G: Per-capita GHG incl. LUC by country group (OWID, 1850–2024)
#
# Groups match Fig C: USA, EU, Other_AII, Other_AI, CHN, IND, Other_nonAI.
# Per-capita = group total_ghg (Mt CO2e) * 1e6 / group population  (t CO2e per person)
# Derived groups are computed by subtraction of underlying group sums so that
# the denominator (population) is correctly group-summed, not taken from any
# single-country population figure.
# ===========================================================================
print("\nBuilding fig9_data.csv ...")

# Four measures: GHG incl/excl LULUCF and CO2 incl/excl LULUCF (OWID only).
# Groups match FigC 8-group split. One row per (measure, year).
# Population from raw OWID (the only column not available in owid_long.csv)
owid_g_pop = pd.read_csv(
    os.path.join(RAW_DIR, "owid-co2-data.csv"),
    usecols=["iso_code", "year", "population"],
    low_memory=False,
)
owid_g_pop = owid_g_pop[owid_g_pop["iso_code"].notna()].copy()
owid_g_pop["year"]       = owid_g_pop["year"].astype(int)
owid_g_pop["population"] = pd.to_numeric(owid_g_pop["population"], errors="coerce")
owid_g_pop = _apply_irl_patch(owid_g_pop, _load_irl_population_patch())

# Emissions from pipeline-corrected owid_long.csv (avoids duplicating the NaN fix from 01_owid_extract.R)
owid_g_em = pd.read_csv(os.path.join(OUT_DIR, "owid_long.csv"))
owid_g_em = owid_g_em[owid_g_em["source"] == "OWID"].copy()
owid_g_em = owid_g_em.pivot_table(
    index=["iso_code", "year"], columns="measure", values="value_Mt", aggfunc="first"
).reset_index()
owid_g_em.columns.name = None
owid_g_em = owid_g_em.rename(columns={
    "ghg_incLUC": "total_ghg",
    "ghg_excLUC": "total_ghg_excluding_lucf",
    "co2_excLUC": "co2",
})

owid_g = owid_g_pop.merge(owid_g_em, on=["iso_code", "year"], how="left")

cg_df = pd.read_csv(os.path.join(QA_DIR, "data", "country_groups.csv"),
                    usecols=["iso_code", "annex_1", "annex_2"])
owid_g = owid_g.merge(cg_df, on="iso_code", how="left")
for col in ["annex_1", "annex_2"]:
    owid_g[col] = owid_g[col].fillna(0).astype(int)

g_rows = []
for yr, ydf in owid_g.groupby("year"):

    def _agg_em(mask, em_col):
        sub = ydf[mask]
        # Numerator: joint dropna — only count emissions where population is also known.
        # Denominator: independent — include all countries with population, even if emissions are NaN.
        em  = float(sub.dropna(subset=[em_col, "population"])[em_col].sum())
        po  = float(sub.dropna(subset=["population"])["population"].sum())
        return em, po

    def pc(em, po):
        return em * 1e6 / po if po > 0 and not pd.isna(po) else np.nan

    for em_col, msr_label in [
        ("total_ghg",               "ghg_incLUC"),
        ("total_ghg_excluding_lucf","ghg_excLUC"),
        ("co2",                     "co2_excLUC"),
        ("co2_incLUC",              "co2_incLUC"),
    ]:
        em_world, po_world = _agg_em(ydf["iso_code"].notna(), em_col)
        em_a2,    po_a2    = _agg_em(ydf["annex_2"] == 1,    em_col)
        em_a1,    po_a1    = _agg_em(ydf["annex_1"] == 1,    em_col)

        g_rows.append({
            "measure":     msr_label,
            "year":        yr,
            "annex_2":     pc(em_a2,            po_a2),
            "annex_1":     pc(em_a1,            po_a1),
            "non_annex_1": pc(em_world - em_a1, po_world - po_a1),
            "world":       pc(em_world,         po_world),
        })

figG = pd.DataFrame(g_rows)
figG = figG[figG["year"] >= 1850]
figG.to_csv(os.path.join(CHART_DIR, "fig9_data.csv"), index=False, float_format="%.3f")
print(f"  Written: fig9_data.csv  ({len(figG)} rows, 4 measures × 4 groups)")


# ===========================================================================
# Fig J: Colonial-attributed cumulative shares (1850-2024 dumbbell)
#
# For each country-year, emissions and GMST contributions are reassigned to
# whoever controlled the territory at the time using Carbon Brief's
# territorial-rule coefficients (data/outputs/colonial_attribution_long.csv,
# generated by scripts/07_extract_colonial.py). Coefficients sum to ~1.0 per
# territory-year and 'Independent' weights stay with the territory itself.
#
# Methodology:
#   Emissions  -- value(country, year) * coefficient(country, year, attr_to)
#                 -> aggregate by attributed_to_iso x source x measure x year.
#   GMST       -- cumulative-to-annual via differencing (year-1 baseline = 0),
#                 attribute each annual increment by the year's coefficient,
#                 then re-cumulate per attributed_to_iso x measure.
#                 Linear approximation: see footnote in
#                 docs/colonial_attribution_method.md.
# ===========================================================================
print("\nBuilding fig8_data.csv (colonial-attributed, 1850-2024) ...")

ATTR_PATH = os.path.join(OUT_DIR, "colonial_attribution_long.csv")
if not os.path.exists(ATTR_PATH):
    raise RuntimeError(
        "colonial_attribution_long.csv not found. "
        "Run scripts/07_extract_colonial.py before this script."
    )
attr_df = pd.read_csv(ATTR_PATH)
attr_df["year"] = attr_df["year"].astype(int)

country_groups_df = pd.read_csv(os.path.join(QA_DIR, "data", "country_groups.csv"))


def reassign_country_values(df, value_col, key_cols):
    """Reassign country-year values via the colonial attribution table.

    df         -- DataFrame with at least iso_code, year, <value_col>
                  (plus any cols listed in key_cols, e.g. source, measure).
    value_col  -- column whose values get split by the attribution coefficient.
    key_cols   -- additional grouping cols carried through the aggregation.

    Countries absent from the attribution table fall through as 100%
    Independent (their values stay with themselves).
    Returns: DataFrame with iso_code = attributed_to_iso, summed by
             (iso_code, *key_cols, year).
    """
    merged = df.merge(
        attr_df,
        left_on=["iso_code", "year"],
        right_on=["territory_iso", "year"],
        how="left",
    )
    fallthrough = merged["coefficient"].isna()
    merged.loc[fallthrough, "coefficient"]       = 1.0
    merged.loc[fallthrough, "attributed_to_iso"] = merged.loc[fallthrough, "iso_code"]

    merged["_attr_value"] = merged[value_col] * merged["coefficient"]

    out = (
        merged.groupby(["attributed_to_iso", *key_cols, "year"], as_index=False)
              ["_attr_value"].sum()
              .rename(columns={"attributed_to_iso": "iso_code",
                               "_attr_value": value_col})
    )
    return out


# --- 1. Reassign country emissions (countries_annual) ---
em_in = cntry[(cntry["year"] >= 1850) & (cntry["year"] <= 2024)][
    ["iso_code", "source", "measure", "year", "value_Mt"]
].copy()
em_J = reassign_country_values(em_in, "value_Mt", ["source", "measure"])

# --- 2. Reassign GMST via differenced annual increments ---
print("  Reassigning GMST via annual-increment differencing ...")
gmst_long = pd.read_csv(os.path.join(OUT_DIR, "gmst_long.csv"))
gmst_long["year"] = gmst_long["year"].astype(int)
gmst_long = gmst_long[gmst_long["year"] <= 2024].copy()
gmst_long = gmst_long.sort_values(["iso_code", "measure", "year"])

# Strip the "gmst_" prefix so measure labels match the emissions side
# (co2_excLUC / co2_incLUC / ghg_excLUC / ghg_incLUC).
gmst_long["measure"] = gmst_long["measure"].str.replace("^gmst_", "", regex=True)

# Annual increment = value(year) - value(year-1); for the first year, use the
# value itself (cumulative starts from the preindustrial baseline of 0 deg C).
gmst_long["increment"] = gmst_long.groupby(["iso_code", "measure"])["value_Mt"].diff()
gmst_long["increment"] = gmst_long["increment"].fillna(gmst_long["value_Mt"])

gmst_inc_J = reassign_country_values(
    gmst_long[["iso_code", "measure", "year", "increment"]],
    value_col="increment",
    key_cols=["measure"],
)
gmst_inc_J = gmst_inc_J.sort_values(["iso_code", "measure", "year"])
gmst_inc_J["value_C"] = gmst_inc_J.groupby(["iso_code", "measure"])["increment"].cumsum()
gmst_J = gmst_inc_J[["iso_code", "measure", "year", "value_C"]]

# --- 3. Aggregate emissions to Annex I / Annex II / non_annex_1 ---
em_J_cum = (
    em_J.groupby(["iso_code", "source", "measure"], as_index=False)["value_Mt"]
        .sum()
        .merge(country_groups_df[["iso_code", "annex_1", "annex_2"]],
               on="iso_code", how="inner")
)

def _grp_sum_em(df, mask):
    return (df[mask].groupby(["source", "measure"], as_index=False)["value_Mt"]
                    .sum()
                    .rename(columns={"value_Mt": "group_Mt"}))

annex1_J = _grp_sum_em(em_J_cum, em_J_cum["annex_1"] == 1)
annex2_J = _grp_sum_em(em_J_cum, em_J_cum["annex_2"] == 1)
nona1_J  = _grp_sum_em(em_J_cum, em_J_cum["annex_1"] == 0)

# World = annex_1 + non_annex_1 (project convention; bunkers naturally absent).
world_J = (annex1_J.merge(nona1_J, on=["source", "measure"], how="outer",
                            suffixes=("_a1", "_na1"))
                     .fillna({"group_Mt_a1": 0, "group_Mt_na1": 0}))
world_J["world_Mt"] = world_J["group_Mt_a1"] + world_J["group_Mt_na1"]
world_J = world_J[["source", "measure", "world_Mt"]]

# --- 4. Aggregate GMST cumulative-at-2024 to Annex I / Annex II / non_annex_1 ---
gmst_J_2024 = (
    gmst_J[gmst_J["year"] == 2024]
        .merge(country_groups_df[["iso_code", "annex_1", "annex_2"]],
               on="iso_code", how="inner")
)

def _grp_sum_gmst(df, mask):
    return (df[mask].groupby(["measure"], as_index=False)["value_C"]
                    .sum()
                    .rename(columns={"value_C": "gmst_C"}))

gmst_a1_J  = _grp_sum_gmst(gmst_J_2024, gmst_J_2024["annex_1"] == 1)
gmst_a2_J  = _grp_sum_gmst(gmst_J_2024, gmst_J_2024["annex_2"] == 1)
gmst_na1_J = _grp_sum_gmst(gmst_J_2024, gmst_J_2024["annex_1"] == 0)

gmst_world_J = (gmst_a1_J.merge(gmst_na1_J, on="measure", how="outer",
                                   suffixes=("_a1", "_na1"))
                            .fillna({"gmst_C_a1": 0, "gmst_C_na1": 0}))
gmst_world_J["gmst_world_C"] = gmst_world_J["gmst_C_a1"] + gmst_world_J["gmst_C_na1"]
gmst_world_J = gmst_world_J[["measure", "gmst_world_C"]]

# --- 5. Share lookups ---
def emit_share_J(group_df, source, measure):
    g = group_df[(group_df["source"] == source) & (group_df["measure"] == measure)]
    w = world_J[(world_J["source"] == source) & (world_J["measure"] == measure)]
    if g.empty or w.empty: return np.nan
    g_val = float(g["group_Mt"].iloc[0]); w_val = float(w["world_Mt"].iloc[0])
    if w_val == 0 or pd.isna(w_val): return np.nan
    return g_val / w_val * 100

def gmst_share_J(group_df, measure):
    g = group_df[group_df["measure"] == measure]
    w = gmst_world_J[gmst_world_J["measure"] == measure]
    if g.empty or w.empty: return np.nan
    g_val = float(g["gmst_C"].iloc[0]); w_val = float(w["gmst_world_C"].iloc[0])
    if w_val == 0 or pd.isna(w_val): return np.nan
    return g_val / w_val * 100

# --- 6. Build fig8_data.csv (mirrors fig3_data.csv shape exactly) ---
SOURCES_J = SOURCES_A   # 1850 baseline -> EDGAR (1970+) and CW (1990+) blank, like Fig A.

dumbbell_measures_J = [
    ("co2_excLUC", MEASURE_LABELS["co2_excLUC"]),
    ("ghg_excLUC", MEASURE_LABELS["ghg_excLUC"]),
    ("co2_incLUC", MEASURE_LABELS["co2_incLUC"]),
    ("ghg_incLUC", MEASURE_LABELS["ghg_incLUC"]),
]
annex_groups_J = [
    ("annex_1", annex1_J, gmst_a1_J, "Annex I"),
    ("annex_2", annex2_J, gmst_a2_J, "Annex II"),
]
rows_J = []
y = 1
for idx, (msr, msr_label) in enumerate(dumbbell_measures_J):
    for grp_key, em_grp_df, gmst_grp_df, grp_label in annex_groups_J:
        row = {"y_pos": y, "label": f"{msr_label} — {grp_label}"}
        active_srcs = SOURCES_J.get(msr, [])
        for src_name, col_name in SRC_TO_COL.items():
            row[col_name] = (emit_share_J(em_grp_df, src_name, msr)
                              if src_name in active_srcs else np.nan)
        row["gmst_ref"] = gmst_share_J(gmst_grp_df, msr)
        rows_J.append(row); y += 1
    if idx < len(dumbbell_measures_J) - 1:
        rows_J.append({"y_pos": y, "label": ""}); y += 1

cols_J = ["y_pos", "label"] + ALL_SRC_COLS + ["gmst_ref"]
figJ = pd.DataFrame(rows_J, columns=cols_J)
figJ.to_csv(os.path.join(CHART_DIR, "fig8_data.csv"),
             index=False, float_format="%.2f")
print(f"  Written: fig8_data.csv  ({len(figJ)} rows)")


# ===========================================================================
# Interactive colonial tabs: colonial-attributed annual MtCO2e per source/group,
# wide format, mirroring interactive_[measure].csv exactly so that the
# interactive can swap between as-reported and colonial via a toggle.
# ===========================================================================
print("\nBuilding interactive_[measure]_colonial.csv (colonial-attributed) ...")

# Annual annex aggregation of em_J: (iso, source, measure, year) -> per-group totals
em_J_grouped = em_J.merge(
    country_groups_df[["iso_code", "annex_1", "annex_2"]], on="iso_code", how="inner",
)
em_J_grouped["non_annex_1"] = (em_J_grouped["annex_1"] == 0).astype(int)

for msr in MEASURES:
    sub = em_J_grouped[em_J_grouped["measure"] == msr]
    by_src = {}
    for src in DB_ALL_SOURCES_RAW:
        s_sub = sub[sub["source"] == src]
        for ann_grp, pfx, flag_col in [
            ("annex_2",     "annex2", "annex_2"),
            ("annex_1",     "annex1", "annex_1"),
            ("non_annex_1", "nona1",  "non_annex_1"),
        ]:
            grp_yearly = (
                s_sub[s_sub[flag_col] == 1]
                .groupby("year", as_index=True)["value_Mt"]
                .sum()
            )
            by_src[f"{pfx}_{src}"] = grp_yearly

    all_years = sorted(
        y for y in set().union(*[s.index for s in by_src.values() if not s.empty])
        if 1750 <= y <= 2024
    )
    rows_raw = []
    for y in all_years:
        row_r = {"year": y}
        for col_name, s in by_src.items():
            v = s.get(y, np.nan)
            row_r[col_name] = float(v) if not pd.isna(v) else np.nan
        rows_raw.append(row_r)
    raw_df = pd.DataFrame(rows_raw)
    raw_df.to_csv(os.path.join(CHART_DIR, f"interactive_{msr}_colonial.csv"),
                  index=False, float_format="%.3f")
    print(f"  Written: interactive_{msr}_colonial.csv  ({len(raw_df)} rows, {len(raw_df.columns)} cols)")

# GMST per-measure cumulative °C, colonial-attributed.
# Column layout: year, then a2/a1/world × 4 measures.
print("Building interactive_gmst_colonial.csv (colonial-attributed cumulative GMST) ...")

gmst_J_grouped = gmst_J.merge(
    country_groups_df[["iso_code", "annex_1", "annex_2"]], on="iso_code", how="inner",
)
gmst_J_grouped["non_annex_1"] = (gmst_J_grouped["annex_1"] == 0).astype(int)

# Aggregate cumulative °C per group × measure × year. World = annex_1 + non_annex_1.
def _gmst_grp_year(df, value_col, flag_col, msr):
    return (
        df[(df[flag_col] == 1) & (df["measure"] == msr)]
        .groupby("year", as_index=True)[value_col]
        .sum()
    )

def _write_gmst_country_agg(df, value_col, out_name):
    """Write country-level cumulative-GMST sums (year | a2/a1/world x 4 measures)."""
    cols = {}
    for msr in MEASURES:
        a2  = _gmst_grp_year(df, value_col, "annex_2",     msr)
        a1  = _gmst_grp_year(df, value_col, "annex_1",     msr)
        na1 = _gmst_grp_year(df, value_col, "non_annex_1", msr)
        cols[f"annex2_{msr}"] = a2
        cols[f"annex1_{msr}"] = a1
        cols[f"world_{msr}"]  = a1.add(na1, fill_value=0)
    yrs = sorted(set().union(*[s.index for s in cols.values() if not s.empty]))
    rows = [{"year": y, **{
        c: (float(s.at[y]) if y in s.index and not pd.isna(s.at[y]) else np.nan)
        for c, s in cols.items()
    }} for y in yrs]
    pd.DataFrame(rows).to_csv(os.path.join(CHART_DIR, out_name),
                                index=False, float_format="%.6f")
    print(f"  Written: {out_name}  ({len(rows)} rows, {len(cols)+1} cols)")

# Colonial-attributed cumulative GMST.
_write_gmst_country_agg(gmst_J_grouped, "value_C", "interactive_gmst_colonial.csv")

# As-reported cumulative GMST aggregated the same way (country-level -> groups).
# Used by the interactive's as-reported path so the colonial toggle changes only
# the coefficients, not the GMST aggregation method. Country-level aggregation
# differs from Jones et al.'s pre-aggregated GLOBAL/ANNEXII rows by ~0.5-1pp;
# the difference is intrinsic to the Jones data.
gmst_long_g = gmst_long.merge(
    country_groups_df[["iso_code", "annex_1", "annex_2"]], on="iso_code", how="inner",
)
gmst_long_g["non_annex_1"] = (gmst_long_g["annex_1"] == 0).astype(int)
_write_gmst_country_agg(gmst_long_g, "value_Mt", "interactive_gmst.csv")


# ===========================================================================
# Interactive raw data: wide-format annual MtCO2e per source/group — inputs
# to script 12 (the interactive dumbbell's payload builder).
# One CSV per measure; columns: year, annex2_[src], annex1_[src], nona1_[src]
# for all 8 sources. NaN where a source has no data for that measure/year.
# ===========================================================================
print("\nBuilding interactive_[measure].csv ...")

for msr in MEASURES:
    ann_series = {}
    for src in DB_ALL_SOURCES_RAW:
        for ann_grp, pfx in [("annex_2", "annex2"), ("annex_1", "annex1"),
                              ("non_annex_1", "nona1")]:
            col = f"{pfx}_{src}"
            ann_series[col] = (
                grp[(grp["source"] == src) & (grp["group"] == ann_grp) &
                    (grp["measure"] == msr)]
                .set_index("year")["value_Mt"]
            )
    all_years_raw = sorted(
        y for y in set().union(*[s.index for s in ann_series.values()])
        if 1750 <= y <= 2024
    )
    rows_raw = []
    for y in all_years_raw:
        row_r = {"year": y}
        for col_name, s in ann_series.items():
            v = s.get(y, np.nan)
            row_r[col_name] = float(v) if not pd.isna(v) else np.nan
        rows_raw.append(row_r)
    raw_df = pd.DataFrame(rows_raw)
    raw_df.to_csv(os.path.join(CHART_DIR, f"interactive_{msr}.csv"),
                  index=False, float_format="%.3f")
    print(f"  Written: interactive_{msr}.csv  ({len(raw_df)} rows, {len(raw_df.columns)} cols)")

# ===========================================================================
# Fig H: Cumulative per-capita emissions as % of world average
#
# Three methods (CGDev Approaches A/B/C):
#   A (EN)  — endpoint-normalised:   cum_emissions / end-year population
#   B (TS)  — trajectory-summed:     sum of annual per-capita figures
#   C (PYW) — person-year-weighted:  cum_emissions / cumulative population
#
# Groups: annex_1, annex_2, world.  Values expressed as % of world average.
# Source: owid_long.csv (pipeline-corrected co2_incLUC, not raw co2_including_luc).
#
# Outputs: fig10_percapita_cumulative_1850.csv, fig11_percapita_cumulative_1990.csv
# ===========================================================================

def build_figH(start_yr, end_yr=2024):
    """Build Fig H cumulative per-capita CSV for a given baseline year.

    Reads owid_long.csv (pipeline-corrected co2_incLUC) and raw OWID population.
    Computes Methods A / B / C per CGDev methodology; expresses each group's
    per-capita figure as a % of the World average.

    Method B group aggregation = sum of country-level TS values (additive by
    definition; see docs/cumulative_per_capita_methods.md §Group aggregation).
    """
    _fignum = {1850: "fig10", 1990: "fig11"}
    print(f"\nBuilding {_fignum[start_yr]}_data.csv ...")

    # --- Emissions: owid_long.csv (co2_incLUC is pipeline-corrected) ---
    ol = pd.read_csv(os.path.join(OUT_DIR, "owid_long.csv"), low_memory=False)
    ol["year"]     = ol["year"].astype(int)
    ol["value_Mt"] = pd.to_numeric(ol["value_Mt"], errors="coerce")

    # Countries only: 3-char ISO, not OWID-prefix aggregates
    ol = ol[
        ol["measure"].isin(MEASURES) &
        ol["iso_code"].notna() &
        (ol["iso_code"].str.len() == 3) &
        (~ol["iso_code"].str.startswith("OWID"))
    ].copy()

    # --- Population: raw OWID (Kosovo has no ISO code in GCB 2025; patch to XKX) ---
    pop = pd.read_csv(
        os.path.join(RAW_DIR, "owid-co2-data.csv"),
        usecols=["country", "iso_code", "year", "population"],
        low_memory=False,
    )
    pop.loc[pop["country"] == "Kosovo", "iso_code"] = "XKX"
    pop["year"]       = pop["year"].astype(int)
    pop["population"] = pd.to_numeric(pop["population"], errors="coerce")
    pop = pop[pop["iso_code"].notna() & (pop["iso_code"].str.len() == 3)].copy()
    pop = _apply_irl_patch(pop, _load_irl_population_patch())

    # Full cumulative population per country for PYW: all years in [start_yr, end_yr]
    # with non-null population, regardless of whether emissions data exists.
    # Missing emission years are treated as 0 (matching original xlsx methodology).
    pop_window = pop[
        (pop["year"] >= start_yr) & (pop["year"] <= end_yr) & pop["population"].notna()
    ]
    full_cum_pop = pop_window.groupby("iso_code")["population"].sum()
    full_end_pop = (
        pop_window.sort_values("year")
        .groupby("iso_code")["population"].last()
    )

    # --- Country group flags ---
    cg = pd.read_csv(
        os.path.join(QA_DIR, "data", "country_groups.csv"),
        usecols=["iso_code", "annex_1", "annex_2"],
    )

    # --- Join population and group flags onto emissions ---
    ol = ol.merge(pop[["iso_code", "year", "population"]], on=["iso_code", "year"], how="left")
    ol = ol.merge(cg,  on="iso_code",           how="left")
    ol["annex_1"] = ol["annex_1"].fillna(0).astype(int)
    ol["annex_2"] = ol["annex_2"].fillna(0).astype(int)

    # --- Year window ---
    ol = ol[(ol["year"] >= start_yr) & (ol["year"] <= end_yr)].copy()

    # --- Country-level aggregation ---
    # Each method is computed per-country first, then summed across groups and world.
    # group_pct = group_sum / world_sum × 100 (additive; Annex I + non-Annex I = 100%).
    #
    # Matching column order in the original xlsx (Cumulative per capita for charts.xlsx):
    #   col A (method_a) = TS  — trajectory-summed: sum of annual per-capita values
    #   col B (method_b) = EN  — endpoint-normalised: cumulative emissions / end-year pop
    #   col C (method_c) = PYW — person-year-weighted: cumulative emissions / cumulative pop
    #
    # PYW cum_pop = sum of all population years in [start_yr, end_yr] for that country,
    # not just years with emissions (missing years treated as 0 emissions).
    c_rows = []
    for (iso, msr), gdf in ol.groupby(["iso_code", "measure"]):
        em_valid = gdf[gdf["value_Mt"].notna()]
        if em_valid.empty:
            continue

        cum_em = float(em_valid["value_Mt"].sum())

        end_pop = float(full_end_pop.get(iso, np.nan))
        cum_pop = float(full_cum_pop.get(iso, np.nan))

        # TS per country: sum of annual (value_Mt * 1e6 / population) where both valid and pop > 0
        ts_rows = gdf[
            gdf["value_Mt"].notna() &
            gdf["population"].notna() &
            (gdf["population"] > 0)
        ].copy()
        val_ts = float((ts_rows["value_Mt"] * 1e6 / ts_rows["population"]).sum()) if not ts_rows.empty else np.nan

        # EN per country: cum_em * 1e6 / end_pop
        val_en = cum_em * 1e6 / end_pop if (pd.notna(end_pop) and end_pop > 0) else np.nan

        # PYW per country: cum_em * 1e6 / cum_pop (full period population, 0 for missing em years)
        val_pyw = cum_em * 1e6 / cum_pop if (pd.notna(cum_pop) and cum_pop > 0) else np.nan

        c_rows.append({
            "iso_code": iso,
            "measure":  msr,
            "cum_em":   cum_em,
            "val_ts":   val_ts,
            "val_en":   val_en,
            "val_pyw":  val_pyw,
            "annex_1":  int(gdf["annex_1"].iloc[0]),
            "annex_2":  int(gdf["annex_2"].iloc[0]),
        })

    c_agg = pd.DataFrame(c_rows)

    # --- Group aggregation: sum country-level method values ---
    def group_agg(subset):
        return subset.groupby("measure", as_index=False).agg(
            cum_em  = ("cum_em",  "sum"),
            val_ts  = ("val_ts",  "sum"),
            val_en  = ("val_en",  "sum"),
            val_pyw = ("val_pyw", "sum"),
        )

    a1 = group_agg(c_agg[c_agg["annex_1"] == 1])
    a2 = group_agg(c_agg[c_agg["annex_2"] == 1])
    ww = group_agg(c_agg)  # all countries = world

    # --- Methods A/B/C as % of world total (group sum / world sum × 100) ---
    # A = TS, B = EN, C = PYW — matches original xlsx column order.
    # abs_share = cumulative absolute Mt share (no per-capita, OWID only).
    def pct_of_world(grp_df, msr):
        g_row = grp_df[grp_df["measure"] == msr]
        w_row = ww[ww["measure"] == msr]
        if g_row.empty or w_row.empty:
            return np.nan, np.nan, np.nan, np.nan
        g = g_row.iloc[0]
        w = w_row.iloc[0]

        a_pct   = g["val_ts"]  / w["val_ts"]  * 100 if (pd.notna(w["val_ts"])  and w["val_ts"]  != 0) else np.nan
        b_pct   = g["val_en"]  / w["val_en"]  * 100 if (pd.notna(w["val_en"])  and w["val_en"]  != 0) else np.nan
        c_pct   = g["val_pyw"] / w["val_pyw"] * 100 if (pd.notna(w["val_pyw"]) and w["val_pyw"] != 0) else np.nan
        abs_pct = g["cum_em"]  / w["cum_em"]  * 100 if (pd.notna(w["cum_em"])  and w["cum_em"]  != 0) else np.nan

        return a_pct, b_pct, c_pct, abs_pct

    # --- Build output: 11 rows (8 data + 3 spacers) ---
    measure_order_H = [
        ("co2_excLUC", "CO2 excl. LULUCF"),
        ("ghg_excLUC", "All GHG excl. LULUCF"),
        ("co2_incLUC", "CO2 incl. LULUCF"),
        ("ghg_incLUC", "All GHG incl. LULUCF"),
    ]

    out_rows = []
    y = 1
    for idx, (msr, msr_label) in enumerate(measure_order_H):
        for grp_label, grp_df in [("Annex I", a1), ("Annex II", a2)]:
            a_pct, b_pct, c_pct, abs_pct = pct_of_world(grp_df, msr)
            out_rows.append({
                "y_pos":      y,
                "label":      f"{msr_label} — {grp_label}",
                "method_a":   a_pct,
                "method_b":   b_pct,
                "method_c":   c_pct,
                "abs_share":  abs_pct,
            })
            y += 1
        if idx < len(measure_order_H) - 1:
            out_rows.append({
                "y_pos": y, "label": "",
                "method_a": np.nan, "method_b": np.nan, "method_c": np.nan,
                "abs_share": np.nan,
            })
            y += 1

    df = pd.DataFrame(out_rows, columns=["y_pos", "label", "method_a", "method_b", "method_c", "abs_share"])
    out_path = os.path.join(CHART_DIR, f"{_fignum[start_yr]}_data.csv")
    df.to_csv(out_path, index=False, float_format="%.2f")
    print(f"  Written: {_fignum[start_yr]}_data.csv  ({len(df)} rows, source=owid_long.csv)")
    return df


print("\nBuilding Fig H data (cumulative per-capita, both baselines) ...")
build_figH(1850)
build_figH(1990)


print("\nDone.")
