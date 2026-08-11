"""
02_extract_climatewatch.py
Extract Climate Watch data for all four gas/LUC measure combinations.

Input:  CW_HistoricalEmissions_ClimateWatch.csv  (QA root)
Output: data/outputs/climatewatch_long.csv

Standard schema: source, iso_code, year, measure, value_Mt

Coverage: 1990-2023 (Climate Watch dataset range, updated 20 Apr 2026)

Sector filter notes:
  - The file uses "LULUCF" in sector names (older downloads used "LUCF" — updated 20 Apr 2026)
  - Some sector names have trailing spaces; strip before filtering
  - The "Country" column contains ISO3 codes (e.g. "AFG")
  - "WORLD" row is excluded by length filter; "EUU" (European Union bloc) is
    explicitly excluded so it doesn't double-count individual EU member rows.

Measures extracted:
  co2_excLUC : Gas="CO2",     Sector="Total excluding LULUCF"
  co2_incLUC : Gas="CO2",     Sector="Total including LULUCF"
  ghg_excLUC : Gas="All GHG", Sector="Total excluding LULUCF"
  ghg_incLUC : Gas="All GHG", Sector="Total including LULUCF"

Units: values in the CSV are already in MtCO2e — no conversion needed.
"""

import os
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW_DIR     = os.path.join(QA_DIR, "data", "raw")
RAW         = os.path.join(RAW_DIR, "CW_HistoricalEmissions_ClimateWatch.csv")
OUT         = os.path.join(QA_DIR, "data", "outputs", "climatewatch_long.csv")

# Measure definitions: (Gas, Sector) -> measure label
MEASURES = {
    ("CO2",     "Total excluding LULUCF"): "co2_excLUC",
    ("CO2",     "Total including LULUCF"): "co2_incLUC",
    ("All GHG", "Total excluding LULUCF"): "ghg_excLUC",
    ("All GHG", "Total including LULUCF"): "ghg_incLUC",
}

print(f"Reading {os.path.basename(RAW)} ...")
df = pd.read_csv(RAW, low_memory=False)

# Strip whitespace from Sector (some values have trailing spaces)
df["Sector"] = df["Sector"].str.strip()

year_cols = [c for c in df.columns if str(c).isdigit()]

chunks = []
for (gas, sector), measure in MEASURES.items():
    subset = df[(df["Gas"] == gas) & (df["Sector"] == sector)].copy()
    if subset.empty:
        print(f"  WARNING: no rows for Gas={gas!r}, Sector={sector!r}")
        continue
    long = subset.melt(
        id_vars=["Country"],
        value_vars=year_cols,
        var_name="year",
        value_name="value_Mt"
    )
    long["year"]     = long["year"].astype(int)
    long["source"]   = "ClimateWatch"
    long["measure"]  = measure
    long["iso_code"] = long["Country"]
    chunks.append(long[["source", "iso_code", "year", "measure", "value_Mt"]])

result = pd.concat(chunks, ignore_index=True)

AGGREGATE_CODES = {"EUU"}   # 3-letter codes that are aggregates, not countries

# Keep only 3-character ISO codes (excludes "WORLD" and blank rows)
result = result[result["iso_code"].str.len() == 3]
# Drop known aggregate codes that share the 3-char shape but represent blocs
result = result[~result["iso_code"].isin(AGGREGATE_CODES)]
result = result.dropna(subset=["value_Mt"])

os.makedirs(os.path.dirname(OUT), exist_ok=True)
result.to_csv(OUT, index=False)
print(
    f"  Written: climatewatch_long.csv  "
    f"({len(result):,} rows, {result['iso_code'].nunique()} countries, "
    f"measures: {sorted(result['measure'].unique())})"
)
