"""
08_stack_and_aggregate.py
Stack all source long-format CSVs and aggregate to group and country levels.

Inputs (from data/outputs/):
  owid_long.csv
  climatwatch_long.csv
  primap_long.csv
  edgar_long.csv
  gcp_long.csv
  gmst_long.csv
  ../data/country_groups.csv

Outputs (to data/outputs/):
  all_sources_stacked.csv  — all country-year-source-measure rows joined with
                              group membership flags (for reference/QA)
  groups_summary.csv       — one row per group x source x measure x year
                              (powers the Summary tab in the workbook)
  countries_annual.csv     — one row per country x source x measure x year
                              (powers the Country tab in the workbook)

Group definitions: every column in country_groups.csv after iso_code and
country_name is treated as a binary group flag.  Any country with a 1 in
that column is included in the group's aggregate.

A special "World" group aggregates all countries regardless of flags.

Note on GMST: values are in degrees Celsius (not MtCO2e).  They are included
in the stacked file and country-level aggregates with measure="gmst_C".
The "World" GMST total is the SUM of country contributions (= global mean
surface temperature response from all countries combined).
"""

import os
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
OUT_DIR     = os.path.join(QA_DIR, "data", "outputs")

SOURCE_FILES = [
    "owid_long.csv",
    "climatewatch_long.csv",
    "primap_long.csv",
    "edgar_long.csv",
    "gcp_long.csv",
    "gmst_long.csv",
]

CG_FILE = os.path.join(QA_DIR, "data", "country_groups.csv")

# ---------------------------------------------------------------------------
# 1. Stack all source files
# ---------------------------------------------------------------------------
print("Stacking source files ...")
chunks = []
for fname in SOURCE_FILES:
    path = os.path.join(OUT_DIR, fname)
    if not os.path.exists(path):
        print(f"  MISSING (skipping): {fname}")
        continue
    df = pd.read_csv(path, low_memory=False)
    # Ensure standard types
    df["year"]     = pd.to_numeric(df["year"], errors="coerce").astype("Int64")
    df["value_Mt"] = pd.to_numeric(df["value_Mt"], errors="coerce")
    chunks.append(df)
    print(f"  Loaded {fname}: {len(df):,} rows")

stacked = pd.concat(chunks, ignore_index=True)
print(f"  Total stacked: {len(stacked):,} rows")

# ---------------------------------------------------------------------------
# 2. Join group membership
# ---------------------------------------------------------------------------
print("Loading country groups ...")
cg = pd.read_csv(CG_FILE)
group_cols = [c for c in cg.columns if c not in ("iso_code", "country_name")]

stacked = stacked.merge(cg[["iso_code"] + group_cols], on="iso_code", how="left")
# Countries not in country_groups get 0 for all group flags
for col in group_cols:
    stacked[col] = stacked[col].fillna(0).astype(int)

# ---------------------------------------------------------------------------
# 3. Write all_sources_stacked.csv (full data with group flags, for QA)
# ---------------------------------------------------------------------------
stacked_path = os.path.join(OUT_DIR, "all_sources_stacked.csv")
stacked.to_csv(stacked_path, index=False)
print(f"  Written: all_sources_stacked.csv ({len(stacked):,} rows)")

# ---------------------------------------------------------------------------
# 4. countries_annual.csv — one row per iso_code x source x measure x year
#    (sum across any within-country duplicates, though there should be none)
# ---------------------------------------------------------------------------
print("Building countries_annual ...")
countries = (
    stacked
    .groupby(["iso_code", "source", "measure", "year"], as_index=False)
    ["value_Mt"].sum()
)
# Add country name for readability
countries = countries.merge(cg[["iso_code", "country_name"]], on="iso_code", how="left")
countries = countries[["iso_code", "country_name", "source", "measure", "year", "value_Mt"]]
countries = countries.sort_values(["iso_code", "source", "measure", "year"])

countries_path = os.path.join(OUT_DIR, "countries_annual.csv")
countries.to_csv(countries_path, index=False)
print(
    f"  Written: countries_annual.csv  "
    f"({len(countries):,} rows, {countries['iso_code'].nunique()} countries)"
)

# ---------------------------------------------------------------------------
# 5. groups_summary.csv — one row per group x source x measure x year
#    For each group column (e.g. annex_1), sum value_Mt for rows where flag=1.
#    Also include a "World" group (all countries).
# ---------------------------------------------------------------------------
print("Building groups_summary ...")

id_cols = ["source", "measure", "year"]

group_chunks = []

# World: all countries
world = stacked.groupby(id_cols, as_index=False)["value_Mt"].sum()
world["group"] = "World"
group_chunks.append(world)

# Each binary group
for gcol in group_cols:
    sub = stacked[stacked[gcol] == 1].groupby(id_cols, as_index=False)["value_Mt"].sum()
    sub["group"] = gcol
    group_chunks.append(sub)

# non_annex_1: countries where annex_1 == 0.
# Kept separate from the loop so it appears explicitly alongside annex_1 in the output,
# making the relationship world = annex_1 + non_annex_1 verifiable.
non_ann1 = stacked[stacked["annex_1"] == 0].groupby(id_cols, as_index=False)["value_Mt"].sum()
non_ann1["group"] = "non_annex_1"
group_chunks.append(non_ann1)

# eu_annex2: EU member states that are also Annex 2 (14 countries: the original EU/EEA
# developed-country bloc that signed Annex 2 of the UNFCCC).
# Distinct from the 'eu' group (27 members) and 'annex_2_eu' column (all Annex 2, not just EU).
eu_a2 = stacked[(stacked["eu"] == 1) & (stacked["annex_2"] == 1)].groupby(
    id_cols, as_index=False)["value_Mt"].sum()
eu_a2["group"] = "eu_annex2"
group_chunks.append(eu_a2)

groups = pd.concat(group_chunks, ignore_index=True)
groups = groups[["group", "source", "measure", "year", "value_Mt"]]
groups = groups.sort_values(["group", "source", "measure", "year"])

groups_path = os.path.join(OUT_DIR, "groups_summary.csv")
groups.to_csv(groups_path, index=False)
print(
    f"  Written: groups_summary.csv  "
    f"({len(groups):,} rows, {groups['group'].nunique()} groups)"
)

print("Done.")
