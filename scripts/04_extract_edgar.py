"""
04_extract_edgar.py
Extract all four EDGAR measures from EDGAR_2025_GHG_booklet_2025.xlsx.

Inputs (single file, two sheets):
  GHG_by_sector_and_country  -> co2_excLUC, ghg_excLUC   (1970-2024)
  LULUCF_countries           -> LULUCF increments         (1990-2024)

Output: data/outputs/edgar_long.csv

Standard schema: source, iso_code, year, measure, value_Mt

Measures produced:
  co2_excLUC  - CO2 fossil+industry only
                (Substance == 'CO2', all 8 sectors from GHG_by_sector_and_country)
  ghg_excLUC  - All GHGs fossil+industry only
                (all substances, all 8 sectors from GHG_by_sector_and_country)
  co2_incLUC  - co2_excLUC + CO2 LULUCF
                (+ Substance == 'CO2', all LULUCF sectors from LULUCF_countries)
  ghg_incLUC  - ghg_excLUC + GHG LULUCF
                (+ CO2 + GWP_100_AR5_CH4 + GWP_100_AR5_N2O, all LULUCF sectors)

Notes:
  - Values in the booklet are already in Mt CO2eq/yr; no unit conversion needed.
  - International aviation (AIR), shipping (SEA), and other non-country aggregates
    are excluded.
  - LULUCF_countries contains a pre-aggregated 'Sector=Net' row per country; this
    is ignored. Totals are computed by summing the component substance rows,
    consistent with the CO2-only extraction.
  - co2_incLUC and ghg_incLUC cover 1990-2024 only (inner join with LULUCF_countries;
    no backfill for pre-1990 years where LULUCF data is absent).

Previous inputs IEA_EDGAR_CO2_1970_2024.xlsx and EDGAR_AR5_GHG_1970_2024.xlsx
are superseded by this script.
"""

import os
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW_DIR     = os.path.join(QA_DIR, "data", "raw")
BOOKLET = os.path.join(RAW_DIR, "EDGAR_2025_GHG_booklet_2025.xlsx")
OUT         = os.path.join(QA_DIR, "data", "outputs", "edgar_long.csv")

NON_COUNTRY = {"AIR", "SEA", "WORLD", "EU27BX"}

ISO_COL = "EDGAR Country Code"

GHG_LULUCF_SUBSTANCES = {"CO2", "GWP_100_AR5_CH4", "GWP_100_AR5_N2O"}


def _booklet_path():
    try:
        open(BOOKLET, "rb").close()
    except PermissionError:
        raise RuntimeError(
            f"Cannot read {os.path.basename(BOOKLET)} — the file is locked, "
            "usually because it is open in Excel. Close it and re-run."
        )
    return BOOKLET


def load_sheet_long(sheet_name):
    """Load a booklet sheet and return a long DataFrame with columns
    [iso_code, substance, sector, year, value_Mt]."""
    df = pd.read_excel(_booklet_path(), sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]

    year_cols = [c for c in df.columns if str(c).isdigit() and 1900 <= int(c) <= 2100]

    id_cols = [ISO_COL, "Substance", "Sector"]
    df = df[id_cols + year_cols].copy()
    df = df.dropna(subset=[ISO_COL])
    df[ISO_COL] = df[ISO_COL].astype(str).str.strip()

    df = df[~df[ISO_COL].isin(NON_COUNTRY)]
    df = df[df[ISO_COL].str.len() == 3]

    long = df.melt(id_vars=id_cols, var_name="year", value_name="value_Mt")
    long["year"] = long["year"].astype(int)
    long["value_Mt"] = pd.to_numeric(long["value_Mt"], errors="coerce")
    long = long.dropna(subset=["value_Mt"])
    long = long.rename(columns={ISO_COL: "iso_code"})
    return long


print("Extracting EDGAR data from booklet ...")

# ── Step 1: excLUC measures from GHG_by_sector_and_country ──────────────────
print("  Reading GHG_by_sector_and_country ...")
sec = load_sheet_long("GHG_by_sector_and_country")

co2_exc = (
    sec[sec["Substance"] == "CO2"]
    .groupby(["iso_code", "year"], as_index=False)["value_Mt"]
    .sum()
    .assign(source="EDGAR", measure="co2_excLUC")
)

ghg_exc = (
    sec
    .groupby(["iso_code", "year"], as_index=False)["value_Mt"]
    .sum()
    .assign(source="EDGAR", measure="ghg_excLUC")
)

print(
    f"    co2_excLUC: {len(co2_exc):,} rows, "
    f"{co2_exc['iso_code'].nunique()} countries, "
    f"years {co2_exc['year'].min()}-{co2_exc['year'].max()}"
)
print(
    f"    ghg_excLUC: {len(ghg_exc):,} rows, "
    f"{ghg_exc['iso_code'].nunique()} countries, "
    f"years {ghg_exc['year'].min()}-{ghg_exc['year'].max()}"
)

# ── Step 2: LULUCF increments from LULUCF_countries ─────────────────────────
print("  Reading LULUCF_countries ...")
luc = load_sheet_long("LULUCF_countries")
luc = luc[luc["Sector"] != "Net"]   # drop pre-aggregated Net rows; sum from components

co2_luc = (
    luc[luc["Substance"] == "CO2"]
    .groupby(["iso_code", "year"], as_index=False)["value_Mt"]
    .sum()
    .rename(columns={"value_Mt": "lulucf_Mt"})
)

ghg_luc = (
    luc[luc["Substance"].isin(GHG_LULUCF_SUBSTANCES)]
    .groupby(["iso_code", "year"], as_index=False)["value_Mt"]
    .sum()
    .rename(columns={"value_Mt": "lulucf_Mt"})
)

print(
    f"    CO2 LULUCF: {len(co2_luc):,} rows, "
    f"{co2_luc['iso_code'].nunique()} countries, "
    f"years {co2_luc['year'].min()}-{co2_luc['year'].max()}"
)
print(
    f"    GHG LULUCF: {len(ghg_luc):,} rows, "
    f"{ghg_luc['iso_code'].nunique()} countries, "
    f"years {ghg_luc['year'].min()}-{ghg_luc['year'].max()}"
)

# ── Step 3: incLUC = excLUC + LULUCF increment ──────────────────────────────
def make_inc(exc_df, luc_df, measure_name):
    # Inner join: only produce incLUC rows for years where LULUCF data exists (1990-2024).
    # A left join + fillna(0) would silently produce pre-1990 rows where incLUC = excLUC,
    # which is misleading because no LULUCF component is included.
    merged = exc_df.merge(luc_df, on=["iso_code", "year"], how="inner")
    merged["value_Mt"] = merged["value_Mt"] + merged["lulucf_Mt"]
    return merged[["source", "iso_code", "year", "measure", "value_Mt"]].assign(
        measure=measure_name
    )

co2_inc = make_inc(co2_exc, co2_luc, "co2_incLUC")
ghg_inc = make_inc(ghg_exc, ghg_luc, "ghg_incLUC")

print(
    f"    co2_incLUC: {len(co2_inc):,} rows, "
    f"{co2_inc['iso_code'].nunique()} countries"
)
print(
    f"    ghg_incLUC: {len(ghg_inc):,} rows, "
    f"{ghg_inc['iso_code'].nunique()} countries"
)

# ── Write output ─────────────────────────────────────────────────────────────
result = pd.concat(
    [co2_exc[["source","iso_code","year","measure","value_Mt"]],
     ghg_exc[["source","iso_code","year","measure","value_Mt"]],
     co2_inc,
     ghg_inc],
    ignore_index=True
)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
result.to_csv(OUT, index=False)
print(
    f"\n  Written: edgar_long.csv  "
    f"({len(result):,} rows, measures: {sorted(result['measure'].unique())})"
)
