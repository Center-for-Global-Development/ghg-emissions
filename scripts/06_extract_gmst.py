"""
06_extract_gmst.py
Extract GMST (Global Mean Surface Temperature) response data.

Input:  GMST_response_1851-2024.csv  (QA root)
Output: data/outputs/gmst_long.csv

Standard schema: source, iso_code, year, measure, value_Mt
  (value_Mt column contains degrees Celsius for GMST; the name is kept
   consistent with other sources for stacking, but units differ)

Outputs 4 measure labels (one per gas/component combination):
  gmst_co2_excLUC  — Gas=CO[2],  Component=Fossil
  gmst_co2_incLUC  — Gas=CO[2],  Component=Total
  gmst_ghg_excLUC  — Gas=3-GHG,  Component=Fossil
  gmst_ghg_incLUC  — Gas=3-GHG,  Component=Total  (equivalent to old gmst_C)

Country-level rows only: ISO3 codes that are exactly 3 characters and not
group-level codes (ANNEXI, ANNEXII, GLOBAL, etc.).

Coverage: 1851-2024
"""

import os
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW_DIR     = os.path.join(QA_DIR, "data", "raw")
RAW         = os.path.join(RAW_DIR, "GMST_response_1851-2024.csv")
OUT         = os.path.join(QA_DIR, "data", "outputs", "gmst_long.csv")

GROUP_CODES = {
    "ANNEXI", "ANNEXII", "GLOBAL", "NONANNEXI", "NONANNEX1",
    "LDC", "LLDC", "SIDS", "EU27", "G20", "OECD",
    "EIT",  # Economies in Transition — 3-char aggregate, not a country
}

MEASURE_MAP = [
    ("CO[2]", "Fossil", "gmst_co2_excLUC"),
    ("CO[2]", "Total",  "gmst_co2_incLUC"),
    ("3-GHG", "Fossil", "gmst_ghg_excLUC"),
    ("3-GHG", "Total",  "gmst_ghg_incLUC"),
]

print(f"Reading {os.path.basename(RAW)} ...")
df = pd.read_csv(RAW)

chunks = []
for gas, comp, measure_label in MEASURE_MAP:
    sub = df[(df["Gas"] == gas) & (df["Component"] == comp)].copy()
    sub = sub.rename(columns={"ISO3": "iso_code", "Year": "year", "Data": "value_Mt"})
    sub = sub[
        (sub["iso_code"].str.len() == 3) &
        ~sub["iso_code"].isin(GROUP_CODES)
    ]
    sub["source"]  = "GMST"
    sub["measure"] = measure_label
    sub = sub[["source", "iso_code", "year", "measure", "value_Mt"]]
    sub = sub.dropna(subset=["value_Mt"])
    chunks.append(sub)
    print(f"  {measure_label}: {len(sub):,} rows, {sub['iso_code'].nunique()} countries")

result = pd.concat(chunks, ignore_index=True)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
result.to_csv(OUT, index=False)
print(
    f"  Written: gmst_long.csv  "
    f"({len(result):,} rows total, {result['measure'].nunique()} measures, "
    f"years {result['year'].min()}-{result['year'].max()})"
)
