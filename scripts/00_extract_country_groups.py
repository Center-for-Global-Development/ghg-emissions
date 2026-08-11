"""
00_extract_country_groups.py
One-off extraction of country group membership from the Combined workbook.

NOTE — provenance only, cannot run from a clean clone. This script reads two
files that are not published with this repository (an internal legacy workbook,
Combined_dropdown_summary_3.1.xlsm, and annex_binary.csv). It is kept to
document how the tracked data/country_groups.csv was produced. It is not part
of main.py; the pipeline consumes the committed CSV directly.

Outputs: ../data/country_groups.csv  (tracked in git — do not gitignore)

Sources:
  - Countries sheet from Combined_dropdown_summary_3.1.xlsm  (all group columns)
  - annex_binary.csv (authoritative source for annex_1 / annex_2)

Column mapping from the Countries sheet (row 4 = headers, row 5+ = data):
  B  = iso_code         F  = ldc              K  = lic
  C  = country_name     G  = lldc             L  = lmic
                        H  = sids             M  = umic
                        I  = lic_lmic_sids_ldc N = hic
                        P  = annex_1 (overridden by annex_binary.csv)
                        Q  = annex_2 (overridden by annex_binary.csv)
                        R  = annex_2_eu
                        S  = non_annex_1  -> EXCLUDED (negate annex_1 if needed)
                        T  = non_annex_2
                        U  = non_annex_2_eu
                        V  = oecd    W = eu    X = brics
                        Y  = g7      Z = g20
                        AB = wb_eca  AC = wb_n_am  AD = wb_lac
                        AE = wb_eap  AF = wb_sa    AG = wb_mena  AH = wb_ssa
                        AO = post_soviet
"""

import os
import openpyxl
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW_DIR     = os.path.join(QA_DIR, "data", "raw")
COMBINED    = os.path.join(QA_DIR, "Combined_dropdown_summary_3.1.xlsm")
ANNEX_CSV   = os.path.join(RAW_DIR, "annex_binary.csv")
OUT_CSV     = os.path.join(QA_DIR, "data", "country_groups.csv")

# Column index (1-based) → output column name.
# S (col 19) = non_annex_1 is intentionally omitted.
COL_MAP = {
     2: "iso_code",
     3: "country_name",
     6: "ldc",
     7: "lldc",
     8: "sids",
     9: "lic_lmic_sids_ldc",
    11: "lic",
    12: "lmic",
    13: "umic",
    14: "hic",
    16: "annex_1",       # will be overridden
    17: "annex_2",       # will be overridden
    18: "annex_2_eu",
    20: "non_annex_2",
    21: "non_annex_2_eu",
    22: "oecd",
    23: "eu",
    24: "brics",
    25: "g7",
    26: "g20",
    28: "wb_eca",
    29: "wb_n_am",
    30: "wb_lac",
    31: "wb_eap",
    32: "wb_sa",
    33: "wb_mena",
    34: "wb_ssa",
    # Note: a post-soviet column exists beyond col 34 but is embedded in a
    # lookup-table region (cols AN-AQ) and cannot be cleanly extracted here.
    # Add manually to country_groups.csv later if needed.
}

DATA_START_ROW = 5   # first country row (1-indexed)
MAX_COL        = 34  # AH = last clean group column (SSA)

print(f"Reading Countries sheet from {os.path.basename(COMBINED)} ...")
wb = openpyxl.load_workbook(COMBINED, read_only=True, keep_vba=False, data_only=True)
ws = wb["Countries"]

rows = []
for row in ws.iter_rows(min_row=DATA_START_ROW, max_col=MAX_COL, values_only=True):
    iso = row[1]   # col B (0-indexed = 1)
    # Keep only rows where iso_code is a 3-letter string (country rows)
    if not isinstance(iso, str) or len(iso) != 3:
        continue
    record = {}
    for col_1based, colname in COL_MAP.items():
        val = row[col_1based - 1]
        # Coerce membership flags to int; leave name/code as-is
        if colname not in ("iso_code", "country_name"):
            val = int(val) if val is not None else 0
        record[colname] = val
    rows.append(record)

wb.close()
cg = pd.DataFrame(rows)
print(f"  Extracted {len(cg)} country rows from Countries sheet.")

# --- Override annex_1 / annex_2 with authoritative annex_binary.csv -----------
annex = pd.read_csv(ANNEX_CSV)[["iso_code", "annex_1", "annex_2"]]
annex["annex_1"] = annex["annex_1"].fillna(0).astype(int)
annex["annex_2"] = annex["annex_2"].fillna(0).astype(int)

# Drop the Combined-sourced annex cols before merging
cg = cg.drop(columns=["annex_1", "annex_2"])
cg = cg.merge(annex, on="iso_code", how="left")
cg["annex_1"] = cg["annex_1"].fillna(0).astype(int)
cg["annex_2"] = cg["annex_2"].fillna(0).astype(int)

# Derive annex_2_eu, non_annex_2, non_annex_2_eu programmatically so they stay
# consistent with annex_2 and eu rather than inheriting Combined workbook values.
#   annex_2_eu = Annex II countries OR EU member states
#   non_annex_2 / non_annex_2_eu = complements of the above
cg["annex_2_eu"]    = ((cg["annex_2"] == 1) | (cg["eu"] == 1)).astype(int)
cg["non_annex_2"]   = (cg["annex_2"]    == 0).astype(int)
cg["non_annex_2_eu"] = (cg["annex_2_eu"] == 0).astype(int)

# Reorder: identifiers first, then annex flags, then rest
id_cols    = ["iso_code", "country_name"]
annex_cols = ["annex_1", "annex_2", "annex_2_eu", "non_annex_2", "non_annex_2_eu"]
other_cols = [c for c in cg.columns if c not in id_cols + annex_cols]
cg = cg[id_cols + annex_cols + other_cols]

os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
cg.to_csv(OUT_CSV, index=False)
print(f"  Saved {len(cg)} rows, {len(cg.columns)} columns -> {OUT_CSV}")
print(f"  Columns: {cg.columns.tolist()}")
