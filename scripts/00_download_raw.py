"""
00_download_raw.py
Fetch the raw source files into data/raw/.

Downloads every source that offers a stable direct link, skipping files that
are already present. Three sources cannot be fetched automatically, and all
three are committed in data/raw/, so a fresh clone needs no manual downloads:

  - Global Carbon Budget national emissions, fossil and land-use change
    (ICOS requires a licence-accept click-through)
  - Climate Watch (no stable direct link)

All three are CC BY 4.0, which permits redistribution with attribution.

Note on reproducibility: OWID updates its published files in place, so a fresh
download may not match the vintage used for the committed outputs. The
versions and download dates used are recorded in data/raw/README.md.

Usage: python scripts/00_download_raw.py
"""

import os
import sys

import requests

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DIR = os.path.join(os.path.dirname(SCRIPTS_DIR), "data", "raw")

# (target filename, url, minimum plausible size in bytes)
# Minimum sizes are ~half the size of the file we used — a guard against
# saving an error page, not an exact check.
DIRECT = [
    ("owid-co2-data.csv",
     "https://owid-public.owid.io/data/co2/owid-co2-data.csv",
     15_000_000),
    ("Guetschow_et_al_2025a-PRIMAP-hist_v2.7_final_no_extrap_no_rounding_22-Aug-2025.csv",
     "https://zenodo.org/records/17090760/files/"
     "Guetschow_et_al_2025a-PRIMAP-hist_v2.7_final_no_extrap_no_rounding_22-Aug-2025.csv"
     "?download=1",
     50_000_000),
    ("EDGAR_2025_GHG_booklet_2025.xlsx",
     "https://edgar.jrc.ec.europa.eu/booklet/EDGAR_2025_GHG_booklet_2025.xlsx",
     2_000_000),
    ("GMST_response_1851-2024.csv",
     "https://zenodo.org/records/16640595/files/GMST_response_1851-2024.csv"
     "?download=1",
     14_000_000),
    ("territorial_rule_database_1850_2023.csv",
     "https://raw.githubusercontent.com/carbonbrief/colonial-emissions-data/"
     "main/output-clean/territorial_rule_database_1850_2023.csv",
     1_000_000),
    ("population.csv",
     "https://ourworldindata.org/grapher/population.csv"
     "?v=1&csvType=full&useColumnShortNames=false",
     1_000_000),
    ("mpd2023_web.xlsx",
     "https://dataverse.nl/api/access/datafile/421302",
     2_000_000),
]

MANUAL = [
    ("National_Fossil_Carbon_Emissions_2025_v0.3.xlsx",
     "https://www.icos-cp.eu/impact/science/global-carbon-budget/2025",
     "Already committed in data/raw/ — nothing to do unless you deleted it. "
     "ICOS serves these behind a licence-accept click-through and now offers "
     "v1.0; this analysis used v0.3."),
    ("National_LandUseChange_Carbon_Emissions_2025v0.2.xlsx",
     "https://www.icos-cp.eu/impact/science/global-carbon-budget/2025",
     "Already committed in data/raw/ — nothing to do unless you deleted it. "
     "ICOS now offers v1.0; this analysis used v0.2."),
    ("CW_HistoricalEmissions_ClimateWatch.csv",
     "https://www.climatewatchdata.org/ghg-emissions",
     "Already committed in data/raw/ — nothing to do unless you deleted it. "
     "A fresh explorer export will differ from the April 2026 extract used."),
]


def fetch(name, url, min_size):
    target = os.path.join(RAW_DIR, name)
    if os.path.exists(target):
        print(f"  present, skipping: {name}")
        return True
    print(f"  downloading: {name}")
    try:
        with requests.get(url, stream=True, timeout=300) as r:
            r.raise_for_status()
            tmp = target + ".part"
            with open(tmp, "wb") as f:
                for chunk in r.iter_content(chunk_size=1 << 20):
                    f.write(chunk)
        size = os.path.getsize(tmp)
        if size < min_size:
            os.remove(tmp)
            print(f"    FAILED: got {size:,} bytes, expected at least "
                  f"{min_size:,} — not saved")
            return False
        os.replace(tmp, target)
        print(f"    ok ({size:,} bytes)")
        return True
    except requests.RequestException as e:
        print(f"    FAILED: {e}")
        return False


def main():
    os.makedirs(RAW_DIR, exist_ok=True)
    print(f"Downloading raw sources into {RAW_DIR}\n")

    failed = [name for name, url, min_size in DIRECT
              if not fetch(name, url, min_size)]

    missing_manual = [(name, url, note) for name, url, note in MANUAL
                      if not os.path.exists(os.path.join(RAW_DIR, name))]

    if missing_manual:
        print("\nThese files must be downloaded by hand:")
        for name, url, note in missing_manual:
            print(f"\n  {name}\n    from: {url}\n    {note}")

    if failed:
        print(f"\n{len(failed)} automatic download(s) failed: "
              + ", ".join(failed))

    if failed or missing_manual:
        sys.exit(1)
    print("\nAll raw files present.")


if __name__ == "__main__":
    main()
