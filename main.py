"""
main.py — run the full GHG emissions analysis pipeline.

Usage:
  python main.py                 run steps 1-12 (extracts through interactive payloads)
  python main.py --download      run scripts/00_download_raw.py first
  python main.py --charts        also run the Excel figure scripts (Windows + Excel only)
  python main.py --only 9,10     run only the listed steps
  python main.py --from 8        resume from a step
  python main.py --skip-validate skip step 11

Step numbers match the script numbers in scripts/.

The figure workbooks (--charts) need Windows with Excel installed and are
committed to the repository, so --charts is off by default.

Raw data must be present in data/raw/ first — see data/raw/README.md or run
with --download.
"""

import argparse
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
RAW_DIR = os.path.join(ROOT, "data", "raw")

# (step number, description, script)
STEPS = [
    (1, "OWID extract", "scripts/01_owid_extract.py"),
    (2, "Climate Watch extract", "scripts/02_extract_climatewatch.py"),
    (3, "PRIMAP extract", "scripts/03_extract_primap.py"),
    (4, "EDGAR extract", "scripts/04_extract_edgar.py"),
    (5, "GCP extract", "scripts/05_extract_gcp.py"),
    (6, "GMST extract", "scripts/06_extract_gmst.py"),
    (7, "Colonial attribution coefficients", "scripts/07_extract_colonial.py"),
    (8, "Stack and aggregate", "scripts/08_stack_and_aggregate.py"),
    (9, "Chart data", "scripts/09_prepare_chart_data.py"),
    (10, "Annex C tables", "scripts/10_build_summary_tables.py"),
    (11, "Validate", "scripts/11_validate.py"),
    (12, "Interactive payloads", "scripts/12_prepare_interactive_dumbbell.py"),
]

# Chart scripts (--charts) — optional, Windows + Excel via COM
CHART_SCRIPTS = [f"scripts/charts/fig{i}.ps1" for i in range(1, 12)]

# Raw files every run needs (Climate Watch is committed; the rest are
# downloaded — see data/raw/README.md)
EXPECTED_RAW = [
    "owid-co2-data.csv",
    "CW_HistoricalEmissions_ClimateWatch.csv",
    "Guetschow_et_al_2025a-PRIMAP-hist_v2.7_final_no_extrap_no_rounding_22-Aug-2025.csv",
    "EDGAR_2025_GHG_booklet_2025.xlsx",
    "National_Fossil_Carbon_Emissions_2025_v0.3.xlsx",
    "National_LandUseChange_Carbon_Emissions_2025v0.2.xlsx",
    "GMST_response_1851-2024.csv",
    "territorial_rule_database_1850_2023.csv",
    "population.csv",
    "mpd2023_web.xlsx",
]


def preflight():
    missing = [f for f in EXPECTED_RAW
               if not os.path.exists(os.path.join(RAW_DIR, f))]
    if missing:
        print("Missing raw source files in data/raw/:\n")
        for f in missing:
            print(f"  {f}")
        print("\nRun `python main.py --download` or see data/raw/README.md "
              "for the download table.")
        sys.exit(1)


def run_step(num, desc, script):
    print(f"\n{'=' * 60}\nStep {num}: {desc}  ({script})\n{'=' * 60}")
    t0 = time.time()
    result = subprocess.run([sys.executable, os.path.join(ROOT, script)],
                            cwd=ROOT)
    elapsed = time.time() - t0
    if result.returncode != 0:
        print(f"\nStep {num} FAILED (exit {result.returncode}) "
              f"after {elapsed:.0f}s — stopping.")
        sys.exit(result.returncode)
    print(f"Step {num} done in {elapsed:.0f}s")


def run_charts():
    for script in CHART_SCRIPTS:
        print(f"\n{'=' * 60}\nCharts: {script}\n{'=' * 60}")
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File",
             os.path.join(ROOT, script)], cwd=ROOT)
        if result.returncode != 0:
            print(f"\n{script} FAILED (exit {result.returncode}) — stopping.")
            sys.exit(result.returncode)


def main():
    p = argparse.ArgumentParser(description="Run the GHG emissions pipeline")
    p.add_argument("--download", action="store_true",
                   help="run scripts/00_download_raw.py first")
    p.add_argument("--charts", action="store_true",
                   help="also run the Excel figure scripts (Windows + Excel)")
    p.add_argument("--only", metavar="N,N",
                   help="run only these step numbers, e.g. --only 9,10")
    p.add_argument("--from", dest="from_step", metavar="N",
                   help="resume from this step number")
    p.add_argument("--skip-validate", action="store_true",
                   help="skip step 11 (validation)")
    args = p.parse_args()

    if args.download:
        result = subprocess.run(
            [sys.executable, os.path.join(ROOT, "scripts/00_download_raw.py")],
            cwd=ROOT)
        if result.returncode != 0:
            print("\nDownload step reported missing files — fix before "
                  "running the pipeline.")
            sys.exit(result.returncode)

    preflight()

    steps = STEPS
    if args.only:
        wanted = {int(s) for s in args.only.split(",")}
        steps = [s for s in steps if s[0] in wanted]
    elif args.from_step:
        steps = [s for s in steps if s[0] >= int(args.from_step)]
    if args.skip_validate:
        steps = [s for s in steps if s[0] != 11]

    t0 = time.time()
    for num, desc, script in steps:
        run_step(num, desc, script)

    if args.charts:
        run_charts()

    print(f"\nPipeline complete in {time.time() - t0:.0f}s")


if __name__ == "__main__":
    main()
