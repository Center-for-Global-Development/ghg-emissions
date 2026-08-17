"""
01_owid_extract.py — OWID data extraction
Converted from 01_owid_extract.R (archived to scripts/archive/).

Produces three outputs:

  data/outputs/owid_long.csv
    Long-format annual emissions for the emissions pipeline (08_stack_and_aggregate.py).
    Columns: source, iso_code, year, measure, value_Mt
    Measures: co2_excLUC, co2_incLUC, ghg_excLUC, ghg_incLUC

  data/outputs/owid_annual_for_cumulative.csv
    Wide-format with population and per-capita columns for the cumulative
    per-capita pipeline.

  data/outputs/owid_coverage_summary.csv
    First/last non-NaN year per country per measure, for QA.

Data source: OWID (Global Carbon Budget 2025 + Jones et al. 2025)
Input: data/raw/owid-co2-data.csv
"""

import os
import pandas as pd
import numpy as np

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW_DIR     = os.path.join(QA_DIR, "data", "raw")
OUT_DIR     = os.path.join(QA_DIR, "data", "outputs")
os.makedirs(OUT_DIR, exist_ok=True)

RAW = os.path.join(RAW_DIR, "owid-co2-data.csv")

# --- 1. Load raw OWID data ---------------------------------------------------

print("Reading owid-co2-data.csv...")
owid_raw = pd.read_csv(
    RAW,
    usecols=[
        "country", "year", "iso_code", "population",
        "co2",                            # CO2 excl. LUC (Mt)
        "co2_per_capita",
        "land_use_change_co2",            # LUC CO2 component (Mt)
        "total_ghg",                      # All GHGs incl. LUC (MtCO2e, AR5)
        "ghg_per_capita",
        "total_ghg_excluding_lucf",       # All GHGs excl. LUC (MtCO2e, AR5)
        "ghg_excluding_lucf_per_capita",
    ],
    dtype={"country": str, "iso_code": str},
)
print(f"  Loaded {len(owid_raw):,} rows")

# --- 2. Kosovo fix + filter to countries ------------------------------------
# Kosovo appears with country="Kosovo" but no ISO code in GCB 2025; assign XKX.

owid_raw.loc[
    (owid_raw["country"] == "Kosovo") & (owid_raw["iso_code"].isna()),
    "iso_code"
] = "XKX"
owid_countries = owid_raw[
    owid_raw["iso_code"].notna() & (owid_raw["iso_code"] != "")
].copy()

# --- 3. Compute co2_incLUC --------------------------------------------------
# co2_incLUC = co2 (fill 0) + luc (fill 0); NaN only when both absent.
# OWID's co2_including_luc propagates NaN from co2 even when luc is present.
# This fix recovers ~13,500 country-year rows across 181 countries.

both_absent = owid_countries["co2"].isna() & owid_countries["land_use_change_co2"].isna()
owid_countries["co2_incLUC_Mt"] = np.where(
    both_absent,
    np.nan,
    owid_countries["co2"].fillna(0) + owid_countries["land_use_change_co2"].fillna(0),
)
owid_countries["co2_incLUC_per_capita_t"] = np.where(
    owid_countries["co2_incLUC_Mt"].isna()
    | owid_countries["population"].isna()
    | (owid_countries["population"] == 0),
    np.nan,
    owid_countries["co2_incLUC_Mt"] * 1e6 / owid_countries["population"],
)

# --- 4. Wide-format output for cumulative per-capita pipeline ---------------

owid_wide = (
    owid_countries[[
        "country", "iso_code", "year", "population",
        "co2", "co2_per_capita",
        "co2_incLUC_Mt", "co2_incLUC_per_capita_t",
        "total_ghg", "ghg_per_capita",
        "total_ghg_excluding_lucf", "ghg_excluding_lucf_per_capita",
    ]]
    .rename(columns={
        "co2":                            "co2_excLUC_Mt",
        "co2_per_capita":                 "co2_excLUC_per_capita_t",
        "total_ghg":                      "ghg_incLUC_Mt",
        "ghg_per_capita":                 "ghg_incLUC_per_capita_t",
        "total_ghg_excluding_lucf":       "ghg_excLUC_Mt",
        "ghg_excluding_lucf_per_capita":  "ghg_excLUC_per_capita_t",
    })
    .sort_values(["country", "year"])
)

owid_wide.to_csv(
    os.path.join(OUT_DIR, "owid_annual_for_cumulative.csv"), index=False, na_rep=""
)
print("Written: owid_annual_for_cumulative.csv")

# --- 5. Long-format output for emissions pipeline ---------------------------

owid_long = owid_wide[
    ["iso_code", "year", "co2_excLUC_Mt", "co2_incLUC_Mt", "ghg_excLUC_Mt", "ghg_incLUC_Mt"]
].melt(
    id_vars=["iso_code", "year"],
    value_vars=["co2_excLUC_Mt", "co2_incLUC_Mt", "ghg_excLUC_Mt", "ghg_incLUC_Mt"],
    var_name="measure",
    value_name="value_Mt",
)
owid_long["measure"] = owid_long["measure"].str.replace("_Mt", "", regex=False)
owid_long = owid_long.dropna(subset=["value_Mt"])
owid_long["source"] = "OWID"
owid_long = (
    owid_long[["source", "iso_code", "year", "measure", "value_Mt"]]
    .sort_values(["iso_code", "year", "measure"])
)

owid_long.to_csv(
    os.path.join(OUT_DIR, "owid_long.csv"), index=False, na_rep=""
)
print(
    f"Written: owid_long.csv  "
    f"({len(owid_long):,} rows, {owid_long['iso_code'].nunique()} countries, "
    f"{owid_long['measure'].nunique()} measures)"
)

# --- 6. Coverage summary for QA ---------------------------------------------

MEASURE_COLS = [
    ("population",        "pop_coverage"),
    ("co2_excLUC_Mt",     "co2_excLUC_coverage"),
    ("co2_incLUC_Mt",     "co2_incLUC_coverage"),
    ("ghg_incLUC_Mt",     "ghg_incLUC_coverage"),
    ("ghg_excLUC_Mt",     "ghg_excLUC_coverage"),
]

def coverage_str(group, col):
    valid = group.loc[group[col].notna(), "year"]
    if len(valid) == 0:
        return "NA-NA"
    return f"{valid.min()}-{valid.max()}"

rows = []
for (country, iso_code), grp in owid_wide.groupby(["country", "iso_code"], sort=True):
    row = {"country": country, "iso_code": iso_code}
    for col, label in MEASURE_COLS:
        row[label] = coverage_str(grp, col)
    rows.append(row)

coverage = pd.DataFrame(rows)
coverage.to_csv(
    os.path.join(OUT_DIR, "owid_coverage_summary.csv"), index=False, na_rep=""
)
print("Written: owid_coverage_summary.csv")
