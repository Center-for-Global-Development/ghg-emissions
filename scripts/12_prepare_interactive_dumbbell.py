"""
12_prepare_interactive_dumbbell.py

Build the JSON data payloads for the interactive web dumbbell
(ghg-emissions-interactive/).

Core payload (data.js / data.json — loaded by every embed):
  group-level annual series from the CSVs produced by 09_prepare_chart_data.py
    charts/interactive_[measure].csv           as-reported annual MtCO2e per source x group
    charts/interactive_[measure]_colonial.csv  colonial-attributed equivalent
    charts/interactive_gmst.csv                as-reported cumulative GMST degC per measure x group
    charts/interactive_gmst_colonial.csv       colonial-attributed equivalent
  plus country picker metadata (name + annex/income/region buckets per country)
  from data/country_groups.csv.

Country payloads (countries_[measure].js — lazy-loaded only when the custom
country picker is used, one file per measure):
  per-country annual MtCO2e per source, as-is (data/outputs/countries_annual.csv)
  and colonial-attributed (coefficients from colonial_attribution_long.csv,
  replicating reassign_country_values() in 09_prepare_chart_data.py), plus
  per-country cumulative GMST degC (gmst_long.csv), as-is and colonial
  (colonial GMST via the same increment-differencing approximation as 09).

The web page computes shares client-side:
  share = sum(group or custom selection, start..end) / (sum(annex1) + sum(nona1)) * 100
  gmst  = (cum[end] - cum[start-1]) / (world[end] - world[start-1]) * 100
          (cum[start-1] treated as 0 when start <= 1851)
which mirrors the SUMIFS/VLOOKUP formulas of the discontinued dynamic-date
dumbbell workbook the interactive replaced.

Pre-req: scripts 07 and 09 must have run.
"""

import csv
import json
import os
from collections import defaultdict

QA_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DATA = os.path.join(QA_DIR, "data", "outputs")
CSV_DIR = os.path.join(OUT_DATA, "charts")
OUT_DIR = os.path.join(QA_DIR, "ghg-emissions-interactive")

MEASURES = ["co2_excLUC", "co2_incLUC", "ghg_excLUC", "ghg_incLUC"]
SOURCES = ["OWID", "PRIMAP", "GCP_fossil", "GCP_BLUE", "GCP_OSCAR",
           "GCP_LUCE", "EDGAR", "ClimateWatch"]
GROUPS = ["annex2", "annex1", "nona1"]

YEAR_MIN = 1850   # slider floor; also the start of the colonial attribution table

ATTRIBUTIONS = {
    "asis":     {"raw_suffix": "",          "gmst_file": "interactive_gmst.csv"},
    "colonial": {"raw_suffix": "_colonial", "gmst_file": "interactive_gmst_colonial.csv"},
}

# countries_annual.csv source names -> payload source keys.
# GCP land-use variants arrive as separate source labels; map straight through.
CA_SOURCE_MAP = {
    "OWID": "OWID", "PRIMAP": "PRIMAP", "EDGAR": "EDGAR",
    "ClimateWatch": "ClimateWatch", "GCP_fossil": "GCP_fossil",
    "GCP_BLUE": "GCP_BLUE", "GCP_OSCAR": "GCP_OSCAR", "GCP_LUCE": "GCP_LUCE",
}

WB_REGIONS = [
    ("wb_eca", "Europe & Central Asia"),
    ("wb_n_am", "North America"),
    ("wb_lac", "Latin America & Caribbean"),
    ("wb_eap", "East Asia & Pacific"),
    ("wb_sa", "South Asia"),
    ("wb_mena", "Middle East & North Africa"),
    ("wb_ssa", "Sub-Saharan Africa"),
]
INCOMES = [("hic", "High income"), ("umic", "Upper-middle income"),
           ("lmic", "Lower-middle income"), ("lic", "Low income")]

# Overlapping membership sets offered as "group by" modes in the country picker.
# Unlike annex/income/region these are not mutually exclusive, so each country
# carries a list of the sets it belongs to and can appear under several headings.
# (flag column in country_groups.csv, label shown in the picker, picker mode)
MEMBERSHIP_SETS = [
    ("g20", "G20", "institutional"),
    ("g7", "G7", "institutional"),
    ("eu", "EU27", "institutional"),
    ("oecd", "OECD", "institutional"),
    ("brics", "BRICS", "institutional"),
    ("ldc", "LDC", "development"),
    ("lldc", "LLDC", "development"),
    ("sids", "SIDS", "development"),
]


def read_csv_dict(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def to_float(s):
    return None if s is None or s.strip() == "" else float(s)


def trim_series(years, values_by_key):
    nonnull = [i for i, y in enumerate(years)
               if any(values_by_key[k][i] is not None for k in values_by_key)]
    if not nonnull:
        return None
    lo, hi = nonnull[0], nonnull[-1]
    return years[lo], years[hi], {k: v[lo:hi + 1] for k, v in values_by_key.items()}


def write_js(path, statement, obj):
    text = json.dumps(obj, separators=(",", ":"))
    with open(path, "w", encoding="utf-8") as f:
        f.write(statement + " = " + text + ";\n")
    return len(text)


# ===========================================================================
# Core payload: group-level series (unchanged logic from v1 of this script)
# ===========================================================================

def build_measure(rows):
    years = [int(r["year"]) for r in rows]
    keep = [i for i, y in enumerate(years) if y >= YEAR_MIN]
    rows = [rows[i] for i in keep]
    years = [years[i] for i in keep]
    out = {}
    for src in SOURCES:
        cols = {g: [to_float(r.get(f"{g}_{src}")) for r in rows] for g in GROUPS}
        trimmed = trim_series(years, cols)
        if trimmed is None:
            continue
        start, end, series = trimmed
        out[src] = {
            "start": start,
            "end": end,
            **{g: [None if v is None else round(v, 3) for v in series[g]]
               for g in GROUPS},
        }
    return out


def build_gmst(rows):
    years = [int(r["year"]) for r in rows]
    out = {"start": years[0], "end": years[-1]}
    for msr in MEASURES:
        out[msr] = {
            grp_out: [to_float(r[f"{grp_csv}_{msr}"]) for r in rows]
            for grp_out, grp_csv in
            [("annex2", "annex2"), ("annex1", "annex1"), ("world", "world")]
        }
    return out


# In the emissions data but absent from country_groups.csv. ANT counts in the
# as-is pipeline's non-Annex-I totals (script 07 aggregates all countries) but
# is dropped from the colonial group totals by script 09's inner join — a known
# ~0.08% inconsistency documented in README.md.
EXTRA_COUNTRIES = [
    {"iso": "ANT", "name": "Netherlands Antilles (dissolved 2010)",
     "annex": "na1", "income": "Unclassified", "region": "Latin America & Caribbean",
     "sets": []},
]


def build_country_meta():
    """Picker metadata: iso, name, annex/income/region buckets, membership sets."""
    rows = read_csv_dict(os.path.join(QA_DIR, "data", "country_groups.csv"))
    meta = list(EXTRA_COUNTRIES)
    for r in rows:
        if int(r["annex_2"]) == 1:
            annex = "a2"
        elif int(r["annex_1"]) == 1:
            annex = "a1"
        else:
            annex = "na1"
        income = next((lbl for flag, lbl in INCOMES if int(r[flag]) == 1), "Unclassified")
        region = next((lbl for flag, lbl in WB_REGIONS if int(r[flag]) == 1), "Other")
        sets = [lbl for flag, lbl, _mode in MEMBERSHIP_SETS if int(r[flag]) == 1]
        meta.append({"iso": r["iso_code"], "name": r["country_name"],
                     "annex": annex, "income": income, "region": region,
                     "sets": sets})
    meta.sort(key=lambda m: m["name"])
    return meta


def build_set_groups():
    """Picker "group by" modes built from overlapping membership sets."""
    modes = {}
    for _flag, label, mode in MEMBERSHIP_SETS:
        modes.setdefault(mode, []).append(label)
    return modes


# ===========================================================================
# Country payloads
# ===========================================================================

def load_attribution():
    """(territory_iso, year) -> [(attributed_to_iso, coefficient), ...]"""
    attr = defaultdict(list)
    for r in read_csv_dict(os.path.join(OUT_DATA, "colonial_attribution_long.csv")):
        attr[(r["territory_iso"], int(r["year"]))].append(
            (r["attributed_to_iso"], float(r["coefficient"])))
    return attr


def reassign(values_by_iso_year, attr):
    """Mirror 09_prepare_chart_data.reassign_country_values: country-years
    absent from the attribution table fall through 100% to themselves."""
    out = defaultdict(float)
    for (iso, year), v in values_by_iso_year.items():
        entries = attr.get((iso, year))
        if entries is None:
            out[(iso, year)] += v
        else:
            for to_iso, coeff in entries:
                out[(to_iso, year)] += v * coeff
    return out


def series_dict(values_by_iso_year, y_min, y_max, ndp):
    """{iso: {"s": start_year, "v": [...]}} with leading/trailing gaps trimmed."""
    by_iso = defaultdict(dict)
    for (iso, year), v in values_by_iso_year.items():
        if y_min <= year <= y_max:
            by_iso[iso][year] = v
    out = {}
    for iso, yv in by_iso.items():
        lo, hi = min(yv), max(yv)
        vals = [None if y not in yv else round(yv[y], ndp) for y in range(lo, hi + 1)]
        if not any(vals):
            continue  # all zero/absent after rounding — contributes nothing
        out[iso] = {"s": lo, "v": vals}
    return out


def build_country_payloads(attr):
    """Per measure: emissions per source (asis + colonial) and GMST (both)."""
    # -- emissions --------------------------------------------------------
    print("Reading countries_annual.csv ...")
    em = defaultdict(dict)   # (measure, source) -> {(iso, year): value}
    for r in read_csv_dict(os.path.join(OUT_DATA, "countries_annual.csv")):
        year = int(r["year"])
        if year < YEAR_MIN:
            continue
        src = CA_SOURCE_MAP.get(r["source"])
        if src is None:
            continue
        em[(r["measure"], src)][(r["iso_code"], year)] = float(r["value_Mt"])

    # -- GMST --------------------------------------------------------------
    print("Reading gmst_long.csv ...")
    gmst = defaultdict(dict)  # measure -> {(iso, year): cumulative degC}
    for r in read_csv_dict(os.path.join(OUT_DATA, "gmst_long.csv")):
        year = int(r["year"])
        if year > 2024:
            continue
        msr = r["measure"].replace("gmst_", "", 1)
        gmst[msr][(r["iso_code"], year)] = float(r["value_Mt"])

    payloads = {}
    for msr in MEASURES:
        block = {"asis": {}, "colonial": {}, "gmst": {}}
        for src in SOURCES:
            vals = em.get((msr, src))
            if not vals:
                continue
            block["asis"][src] = series_dict(vals, YEAR_MIN, 2024, 3)
            block["colonial"][src] = series_dict(reassign(vals, attr), YEAR_MIN, 2024, 3)

        # GMST colonial: difference cumulative into annual increments, apply
        # coefficients, re-cumulate (linear approximation — same as script 09,
        # see docs/colonial_attribution_method.md).
        cum = gmst[msr]
        by_iso = defaultdict(dict)
        for (iso, year), v in cum.items():
            by_iso[iso][year] = v
        increments = {}
        for iso, yv in by_iso.items():
            prev = 0.0
            for y in sorted(yv):
                increments[(iso, y)] = yv[y] - prev
                prev = yv[y]
        inc_J = reassign(increments, attr)
        cum_J = {}
        by_iso_J = defaultdict(dict)
        for (iso, year), v in inc_J.items():
            by_iso_J[iso][year] = v
        for iso, yv in by_iso_J.items():
            running = 0.0
            for y in sorted(yv):
                running += yv[y]
                cum_J[(iso, y)] = running

        block["gmst"]["asis"] = series_dict(cum, 1851, 2024, 7)
        block["gmst"]["colonial"] = series_dict(cum_J, 1851, 2024, 7)
        payloads[msr] = block
    return payloads


def main():
    # ---- core payload ----------------------------------------------------
    payload = {"measures": MEASURES, "sources": SOURCES, "attributions": {},
               "countries": build_country_meta(), "setGroups": build_set_groups()}
    for attr_name, spec in ATTRIBUTIONS.items():
        block = {"measures": {}}
        for msr in MEASURES:
            path = os.path.join(CSV_DIR, f"interactive_{msr}{spec['raw_suffix']}.csv")
            block["measures"][msr] = build_measure(read_csv_dict(path))
        block["gmst"] = build_gmst(
            read_csv_dict(os.path.join(CSV_DIR, spec["gmst_file"])))
        payload["attributions"][attr_name] = block

    os.makedirs(OUT_DIR, exist_ok=True)
    core_text = json.dumps(payload, separators=(",", ":"))
    with open(os.path.join(OUT_DIR, "data.json"), "w", encoding="utf-8") as f:
        f.write(core_text)
    with open(os.path.join(OUT_DIR, "data.js"), "w", encoding="utf-8") as f:
        f.write("window.EMISSIONS_DATA = " + core_text + ";\n")
    print(f"Wrote data.js ({len(core_text):,} bytes)")

    # ---- country payloads (lazy-loaded) -----------------------------------
    attr = load_attribution()
    payloads = build_country_payloads(attr)
    for msr, block in payloads.items():
        n = write_js(
            os.path.join(OUT_DIR, f"countries_{msr}.js"),
            'window.EMISSIONS_COUNTRIES = window.EMISSIONS_COUNTRIES || {}; '
            f'window.EMISSIONS_COUNTRIES["{msr}"]',
            block)
        n_iso = len(set().union(*[set(block["asis"][s]) for s in block["asis"]]))
        print(f"Wrote countries_{msr}.js ({n:,} bytes; {n_iso} countries)")

    # ---- coverage report ---------------------------------------------------
    for attr_name, block in payload["attributions"].items():
        print(f"\n[{attr_name}]")
        for msr, srcs in block["measures"].items():
            cov = ", ".join(f"{s} {d['start']}-{d['end']}" for s, d in srcs.items())
            print(f"  {msr}: {cov}")
        g = block["gmst"]
        print(f"  GMST: {g['start']}-{g['end']}")


if __name__ == "__main__":
    main()
