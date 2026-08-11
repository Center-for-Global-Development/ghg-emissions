"""
03_extract_primap.py
Extract PRIMAP-hist v2.7 data for all four gas/LUC measure combinations.

Input:  Guetschow_et_al_2025a-PRIMAP-hist_v2.7_final_no_extrap_no_rounding_22-Aug-2025.csv
Output: data/outputs/primap_long.csv

Standard schema: source, iso_code, year, measure, value_Mt

Coverage: 1750-2024

Filtering logic:
  scenario  = "HISTCR"               (country-reported; main scenario)
  entity    = "CO2"                  for CO2 measures
  entity    = "KYOTOGHG (AR5GWP100)" for All GHG measures (AR5 GWPs, consistent with EDGAR)
  category  = "0"                    total incl. LULUCF
  category  = "M.0.EL"              total excl. LULUCF

Measures extracted:
  co2_excLUC : entity=CO2,                    category=M.0.EL
  co2_incLUC : entity=CO2,                    category=0
  ghg_excLUC : entity=KYOTOGHG (AR5GWP100),   category=M.0.EL
  ghg_incLUC : entity=KYOTOGHG (AR5GWP100),   category=0

Unit conversion: raw unit is "CO2 * gigagram / yr" (Gg = 1000 t)
  value_Mt = raw_value * 0.001  (Gg -> Mt)

Notes:
  - No WORLD or regional aggregates in PRIMAP; all rows are country ISO3 codes.
  - The file is large (~1.4 GB uncompressed); pandas reads it in chunks for speed.
"""

import os
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW_DIR     = os.path.join(QA_DIR, "data", "raw")

# Find the PRIMAP file (name contains version string)
raw_candidates = [f for f in os.listdir(RAW_DIR) if f.startswith("Guetschow") and f.endswith(".csv")]
if not raw_candidates:
    raise FileNotFoundError("PRIMAP CSV not found in data/raw/.")
RAW = os.path.join(RAW_DIR, raw_candidates[0])
OUT = os.path.join(QA_DIR, "data", "outputs", "primap_long.csv")

# Measure definitions: (entity, category) -> measure label
MEASURES = {
    ("CO2",                    "0"):     "co2_incLUC",
    ("CO2",                    "M.0.EL"): "co2_excLUC",
    ("KYOTOGHG (AR5GWP100)",   "0"):     "ghg_incLUC",
    ("KYOTOGHG (AR5GWP100)",   "M.0.EL"): "ghg_excLUC",
}

ENTITY_COL   = "entity"
CATEGORY_COL = "category (IPCC2006_PRIMAP)"
SCENARIO_COL = "scenario (PRIMAP-hist)"
AREA_COL     = "area (ISO3)"

print(f"Reading {os.path.basename(RAW)} ...")
df = pd.read_csv(RAW, low_memory=False)
print(f"  Loaded {len(df):,} rows.")

# Apply filters
mask = (
    (df[SCENARIO_COL] == "HISTCR") &
    (df[ENTITY_COL].isin([e for e, _ in MEASURES.keys()])) &
    (df[CATEGORY_COL].isin([c for _, c in MEASURES.keys()]))
)
df = df[mask].copy()
print(f"  After filtering: {len(df):,} rows.")

# Year columns are numeric strings
year_cols = [c for c in df.columns if str(c).isdigit()]

chunks = []
for (entity, category), measure in MEASURES.items():
    subset = df[(df[ENTITY_COL] == entity) & (df[CATEGORY_COL] == category)].copy()
    if subset.empty:
        print(f"  WARNING: no rows for entity={entity!r}, category={category!r}")
        continue
    long = subset[[AREA_COL] + year_cols].melt(
        id_vars=[AREA_COL],
        var_name="year",
        value_name="raw_value"
    )
    long["year"]     = long["year"].astype(int)
    long["value_Mt"] = long["raw_value"] * 0.001   # Gg -> Mt
    long["source"]   = "PRIMAP"
    long["measure"]  = measure
    long["iso_code"] = long[AREA_COL]
    chunks.append(long[["source", "iso_code", "year", "measure", "value_Mt"]])

result = pd.concat(chunks, ignore_index=True)
result = result.dropna(subset=["value_Mt"])

os.makedirs(os.path.dirname(OUT), exist_ok=True)
result.to_csv(OUT, index=False)
print(
    f"  Written: primap_long.csv  "
    f"({len(result):,} rows, {result['iso_code'].nunique()} countries, "
    f"years {result['year'].min()}-{result['year'].max()}, "
    f"measures: {sorted(result['measure'].unique())})"
)
