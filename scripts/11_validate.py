"""
11_validate.py - Pipeline invariant checks.

Run after re-running 08_stack_and_aggregate.py to catch regressions in the
pipeline outputs. Each check is a simple assertion grounded in a real bug
this project has hit at least once. Exit code 0 = all pass, 1 = any failure.

Usage:
    python scripts/11_validate.py

Checks (most are oneshot lessons from the chart-changelog audit trail):
  1. groups_summary.csv has the expected columns.
  2. groups_summary contains a non_annex_1 group.
  3. World = annex_1 + non_annex_1 in groups_summary (per source, measure, year).
  4. countries_annual.csv contains no known aggregate codes (catches EIT-style
     leaks of group totals masquerading as country rows).
  5. GMST has all four per-measure labels (catches reverting to single gmst_C).
  6. Source x measure coverage matrix matches the documented support table
     (catches e.g. EDGAR appearing under co2_incLUC, GCP_fossil under ghg_*).
  7. Source year ranges match the documented coverage (catches silent truncation
     when an upstream download changes).

Note on the annex_2 partition: World != annex_2 + non_annex_2 in this dataset
because some source rows (e.g. ClimateWatch's "European Union" aggregate)
contribute to World but have no annex_2 flag. The annex_1 partition (check 3)
is exact by construction; annex_2 isn't, so it isn't checked.
"""
from __future__ import annotations

import sys
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "data" / "outputs"
DATA = ROOT / "data"

# Documented source x measure support. True = expected to appear in the stacked
# outputs, False = expected to be absent. See README.md / docs/pipeline-status.md.
EXPECTED_COVERAGE = {
    "OWID":         {"co2_excLUC": True,  "co2_incLUC": True,  "ghg_excLUC": True,  "ghg_incLUC": True},
    "ClimateWatch": {"co2_excLUC": True,  "co2_incLUC": True,  "ghg_excLUC": True,  "ghg_incLUC": True},
    "PRIMAP":       {"co2_excLUC": True,  "co2_incLUC": True,  "ghg_excLUC": True,  "ghg_incLUC": True},
    "EDGAR":        {"co2_excLUC": True,  "co2_incLUC": True,  "ghg_excLUC": True,  "ghg_incLUC": True},
    "GCP_fossil":   {"co2_excLUC": True,  "co2_incLUC": False, "ghg_excLUC": False, "ghg_incLUC": False},
    "GCP_BLUE":     {"co2_excLUC": False, "co2_incLUC": True,  "ghg_excLUC": False, "ghg_incLUC": False},
    "GCP_OSCAR":    {"co2_excLUC": False, "co2_incLUC": True,  "ghg_excLUC": False, "ghg_incLUC": False},
    "GCP_LUCE":     {"co2_excLUC": False, "co2_incLUC": True,  "ghg_excLUC": False, "ghg_incLUC": False},
}

# GMST has its own per-measure labels rather than the four emission measures
# (gmst_co2_excLUC, gmst_co2_incLUC, gmst_ghg_excLUC, gmst_ghg_incLUC).
GMST_MEASURES = {f"gmst_{m}" for m in ("co2_excLUC", "co2_incLUC", "ghg_excLUC", "ghg_incLUC")}

# Documented year ranges per source. (min_year, max_year) inclusive.
# These reflect the current data, not historical doc strings — README.md and
# 05_extract_gcp.py docstrings still say 1750-2024 for GCP_fossil and 1850-2023
# for GCP LUC variants. The actual data is 1850-2024 for all four GCP series
# (GCB 2025 v0.3 / v0.2 update). Update both docs and this dict in lockstep.
EXPECTED_YEARS = {
    "OWID":         (1750, 2024),
    "ClimateWatch": (1990, 2023),
    "PRIMAP":       (1750, 2024),
    "EDGAR":        (1970, 2024),
    "GCP_fossil":   (1850, 2024),
    "GCP_BLUE":     (1850, 2024),
    "GCP_OSCAR":    (1850, 2024),
    "GCP_LUCE":     (1850, 2024),
    "GMST":         (1851, 2024),
}

# Codes that are aggregates / non-country totals — must never appear in
# countries_annual.csv. Pattern lifted from the historical EIT bug
# (06_extract_gmst.py was passing EIT through because it was missing from
# GROUP_CODES). Add to this list when new aggregates surface.
NON_COUNTRY_CODES = {
    "EIT",                       # Economies in Transition (Jones et al. GMST)
    "EUU", "EU27", "EU28",       # European Union bloc variants
    "ANNEX1", "ANNEX2",          # UNFCCC group totals
    "NONANEX1", "NONANNEX1",
    "AOSIS",                     # Alliance of Small Island States
    "LDC", "LLDC", "SIDS",       # development-status aggregates
    "OECD",
    "AFRICA", "ASIA", "EUROPE",
    "WORLD", "GLOBAL",
}

TOL = 1e-3   # 1 kg tolerance for sum-equality checks (values are in MtCO2e)


# ---------------------------------------------------------------------------

results: list[tuple[str, bool, str]] = []


def check(name: str, fn):
    try:
        ok, msg = fn()
    except Exception as e:
        ok, msg = False, f"raised {type(e).__name__}: {e}"
    results.append((name, ok, msg))


# ---------------------------------------------------------------------------

def c1_groups_summary_columns():
    g = pd.read_csv(OUT / "groups_summary.csv", nrows=1)
    expected = {"group", "source", "measure", "year", "value_Mt"}
    missing = expected - set(g.columns)
    extra   = set(g.columns) - expected
    if missing:
        return False, f"missing columns: {sorted(missing)}"
    if extra:
        return False, f"unexpected extra columns: {sorted(extra)}"
    return True, "columns = group, source, measure, year, value_Mt"


def c2_non_annex_1_group_exists():
    g = pd.read_csv(OUT / "groups_summary.csv", usecols=["group"])
    groups = set(g["group"].unique())
    if "non_annex_1" not in groups:
        return False, "non_annex_1 group missing - script 07 regressed"
    return True, f"non_annex_1 present (alongside {len(groups)-1} other groups)"


def _pivot_world_partition(g: pd.DataFrame, a: str, b: str) -> tuple[pd.Series, pd.Series, int]:
    """For each (source, measure, year), return (world, a + b, n_rows)."""
    sub = g[g["group"].isin(["World", a, b])]
    pv = sub.pivot_table(index=["source", "measure", "year"],
                         columns="group", values="value_Mt", aggfunc="sum")
    # Restrict to rows where all three are present (some sources/measures may not
    # have all three groups populated for every year; that's fine).
    pv = pv.dropna(subset=["World", a, b])
    return pv["World"], pv[a] + pv[b], len(pv)


def c3_world_eq_annex_1_plus_non_annex_1():
    g = pd.read_csv(OUT / "groups_summary.csv")
    world, expected, n = _pivot_world_partition(g, "annex_1", "non_annex_1")
    if n == 0:
        return False, "no (source, measure, year) rows with all three groups"
    diff = (world - expected).abs()
    bad = diff[diff > TOL]
    if len(bad) == 0:
        return True, f"{n} rows checked: world == annex_1 + non_annex_1 within {TOL} Mt"
    return False, (f"{len(bad)} of {n} rows fail; max |diff| = {diff.max():.4f} Mt; "
                   f"first bad row: {bad.index[0]}")


def c4_no_aggregate_codes_in_countries_annual():
    ca = pd.read_csv(OUT / "countries_annual.csv", usecols=["iso_code", "source"])
    found = ca[["iso_code", "source"]].drop_duplicates()
    leaked = found[found["iso_code"].isin(NON_COUNTRY_CODES)]
    if len(leaked) == 0:
        return True, f"none of {len(NON_COUNTRY_CODES)} known aggregate codes appear as countries"
    rows = ", ".join(f"{r.iso_code}@{r.source}" for r in leaked.itertuples())
    return False, f"{len(leaked)} aggregate code(s) leaking into countries_annual: {rows}"


def c5_gmst_has_four_per_measure_labels():
    g = pd.read_csv(OUT / "gmst_long.csv", usecols=["measure"])
    found = set(g["measure"].unique())
    missing = GMST_MEASURES - found
    extra = found - GMST_MEASURES
    if missing:
        return False, f"gmst_long.csv missing measure labels: {sorted(missing)}"
    if extra:
        return False, f"gmst_long.csv has unexpected measure labels: {sorted(extra)}"
    return True, f"gmst_long.csv has all 4 per-measure labels: {sorted(GMST_MEASURES)}"


def c6_source_measure_coverage():
    """For each (source, measure) cell in EXPECTED_COVERAGE, the stacked output
    must contain rows iff support is True."""
    g = pd.read_csv(OUT / "groups_summary.csv", usecols=["source", "measure"])
    pairs = set(map(tuple, g.drop_duplicates().values))   # (source, measure) actually present
    failures = []
    for src, msrs in EXPECTED_COVERAGE.items():
        for msr, expected in msrs.items():
            actual = (src, msr) in pairs
            if actual != expected:
                failures.append(f"  {src} x {msr}: expected={'present' if expected else 'absent'}, got={'present' if actual else 'absent'}")
    if not failures:
        return True, f"all {sum(len(m) for m in EXPECTED_COVERAGE.values())} (source, measure) cells match expected"
    return False, f"{len(failures)} mismatches:\n" + "\n".join(failures)


def c7_source_year_ranges():
    g = pd.read_csv(OUT / "groups_summary.csv", usecols=["source", "year"])
    actual_ranges = g.groupby("source")["year"].agg(["min", "max"])
    failures = []
    for src, (exp_lo, exp_hi) in EXPECTED_YEARS.items():
        if src not in actual_ranges.index:
            failures.append(f"  {src}: not present in groups_summary.csv")
            continue
        a_lo, a_hi = int(actual_ranges.loc[src, "min"]), int(actual_ranges.loc[src, "max"])
        if (a_lo, a_hi) != (exp_lo, exp_hi):
            failures.append(f"  {src}: expected {exp_lo}-{exp_hi}, got {a_lo}-{a_hi}")
    if not failures:
        return True, f"all {len(EXPECTED_YEARS)} sources span the expected year range"
    return False, f"{len(failures)} mismatch(es):\n" + "\n".join(failures)


# ---------------------------------------------------------------------------

CHECKS = [
    ("groups_summary columns",                   c1_groups_summary_columns),
    ("non_annex_1 group present",                c2_non_annex_1_group_exists),
    ("World = annex_1 + non_annex_1",            c3_world_eq_annex_1_plus_non_annex_1),
    ("No aggregate codes in countries_annual",   c4_no_aggregate_codes_in_countries_annual),
    ("GMST has 4 per-measure labels",            c5_gmst_has_four_per_measure_labels),
    ("Source x measure coverage matrix",         c6_source_measure_coverage),
    ("Source year ranges",                       c7_source_year_ranges),
]


def main() -> int:
    print(f"Running {len(CHECKS)} pipeline invariant checks...\n")
    for name, fn in CHECKS:
        check(name, fn)

    width = max(len(n) for n, _, _ in results)
    failed = 0
    for name, ok, msg in results:
        marker = "PASS" if ok else "FAIL"
        print(f"  [{marker}] {name.ljust(width)}  {msg}")
        if not ok:
            failed += 1

    print()
    if failed == 0:
        print(f"All {len(results)} checks passed.")
        return 0
    print(f"{failed} of {len(results)} checks FAILED.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
