"""
07_extract_colonial.py
Reshape Carbon Brief's territorial-rule database into long format for use
in the colonial-attribution analysis (Figure 8 and the interactive's toggle).

Input:  data/raw/territorial_rule_database_1850_2023.csv
Output: data/outputs/colonial_attribution_long.csv

Schema of output:
    territory_iso        — ISO3 of the territory whose emissions are being split
    year                 — int (1850-2024; the raw CSV ends 2023, forward-filled to 2024)
    attributed_to_iso    — ISO3 of the entity emissions are reassigned to.
                           For 'Independent' coefficient, equals territory_iso.
    coefficient          — float in (0, 1]; rows with zero coefficient are dropped.

Drops:
    - The Yugoslavia rows (blank ISO3 in the raw CSV). Successor states
      (HRV, SVN, SRB, MKD, BIH, MNE) appear from their independence year.
"""

import os
import pandas as pd

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
RAW         = os.path.join(QA_DIR, "data", "raw",
                            "territorial_rule_database_1850_2023.csv")
OUT_DIR     = os.path.join(QA_DIR, "data", "outputs")
OUT         = os.path.join(OUT_DIR, "colonial_attribution_long.csv")

# Map each colonial-power column header to its modern ISO3 code.
# 'Independent' is special-cased to use the territory's own ISO.
POWER_TO_ISO = {
    "Turkey":         "TUR",
    "Netherlands":    "NLD",
    "Portugal":       "PRT",
    "Spain":          "ESP",
    "France":         "FRA",
    "United Kingdom": "GBR",
    "Germany":        "DEU",
    "Belgium":        "BEL",
    "China":          "CHN",
    "Japan":          "JPN",
    "Italy":          "ITA",
    "USA":            "USA",
    "Australia":      "AUS",
    "Austria":        "AUT",
    "Hungary":        "HUN",
    "Russia":         "RUS",
}

EXTEND_TO_YEAR = 2024   # forward-fill 2023 coefficients to cover 2024


def main():
    print(f"Reading {os.path.basename(RAW)} ...")
    df = pd.read_csv(RAW)

    # Drop blank-ISO rows (Yugoslavia).
    n_before = len(df)
    df = df[df["iso_code"].notna() & (df["iso_code"].str.strip() != "")].copy()
    n_dropped = n_before - len(df)
    if n_dropped:
        print(f"  Dropped {n_dropped} blank-ISO rows (Yugoslavia)")

    df["year"] = df["year"].astype(int)

    power_cols = list(POWER_TO_ISO.keys()) + ["Independent"]
    missing = [c for c in power_cols if c not in df.columns]
    if missing:
        raise RuntimeError(f"Expected columns missing from CSV: {missing}")

    long = df.melt(
        id_vars=["iso_code", "country", "year"],
        value_vars=power_cols,
        var_name="power",
        value_name="coefficient",
    )
    long = long[long["coefficient"] > 0].copy()

    long["attributed_to_iso"] = long.apply(
        lambda r: r["iso_code"] if r["power"] == "Independent"
                  else POWER_TO_ISO[r["power"]],
        axis=1,
    )
    long = long.rename(columns={"iso_code": "territory_iso"})
    long = long[["territory_iso", "year", "attributed_to_iso", "coefficient"]]

    sums = long.groupby(["territory_iso", "year"])["coefficient"].sum()
    bad = sums[(sums - 1.0).abs() > 1e-3]
    if not bad.empty:
        print(f"  WARNING: {len(bad)} (territory, year) coefficient sums deviate from 1.0:")
        print(bad.head(10))

    last_yr = int(long["year"].max())
    if EXTEND_TO_YEAR > last_yr:
        last = long[long["year"] == last_yr].copy()
        extra = []
        for y in range(last_yr + 1, EXTEND_TO_YEAR + 1):
            tmp = last.copy()
            tmp["year"] = y
            extra.append(tmp)
        long = pd.concat([long] + extra, ignore_index=True)
        print(f"  Forward-filled {last_yr}->{EXTEND_TO_YEAR} ({EXTEND_TO_YEAR - last_yr} years)")

    long = long.sort_values(["territory_iso", "year", "attributed_to_iso"]).reset_index(drop=True)

    os.makedirs(OUT_DIR, exist_ok=True)
    long.to_csv(OUT, index=False, float_format="%.6f")
    print(
        f"  Written: {os.path.relpath(OUT, QA_DIR)}  "
        f"({len(long):,} rows, {long['territory_iso'].nunique()} territories, "
        f"years {long['year'].min()}-{long['year'].max()})"
    )


if __name__ == "__main__":
    main()
