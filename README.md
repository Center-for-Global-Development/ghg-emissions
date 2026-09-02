# Comparative GHG emissions analysis

Comparative analysis of national greenhouse-gas emissions across six data
sources — Our World in Data (OWID), Climate Watch, PRIMAP-hist, EDGAR, the
Global Carbon Project (GCP), and the Jones et al. national temperature
contributions (GMST) — supporting the Center for Global Development's climate
finance fair shares work (building on
[Beynon 2024](https://www.cgdev.org/publication/why-climate-finance-fair-shares-model-metrics-matter)).
The pipeline standardises every source to a common schema, aggregates to
UNFCCC country groups, and produces the paper's figures, summary tables, and
an interactive source-comparison chart.

## Data sources

| Dataset | Source | Coverage |
|---------|--------|----------|
| OWID CO2 data | [Our World in Data](https://github.com/owid/co2-data) (Global Carbon Budget 2025) | 1750–2024, CO2 + all GHGs |
| Climate Watch | [Climate Watch Historical GHG Emissions](https://www.climatewatchdata.org/ghg-emissions) (WRI) | 1990–2023, all GHGs |
| PRIMAP-hist v2.7 | [Gütschow et al. (2025)](https://doi.org/10.5281/zenodo.17090760) | 1750–2024, all GHGs |
| EDGAR 2025 booklet | [JRC / IEA EDGAR Community GHG Database](https://edgar.jrc.ec.europa.eu/booklet/EDGAR_2025_GHG_booklet_2025.xlsx) | 1970–2024, CO2 + all GHGs excl. LULUCF; 1990–2024 incl. LULUCF |
| GCP (fossil) | [Global Carbon Project](https://www.icos-cp.eu/impact/science/global-carbon-budget/2025) (GCB 2025) | 1850–2024, fossil CO2 |
| GCP (land-use change) | [Global Carbon Project](https://www.icos-cp.eu/impact/science/global-carbon-budget/2025) (GCB 2025) | 1850–2024, LUC CO2 (3 models: BLUE, OSCAR, LUCE) |
| GMST | [Jones et al., National contributions to climate change](https://doi.org/10.5281/zenodo.16640595) (v2025.1) | 1851–2024, temperature response (°C) |
| Territorial rule database | [Carbon Brief, colonial-emissions-data](https://github.com/carbonbrief/colonial-emissions-data) | 1850–2023, colonial-rule coefficients |

[`data/raw/README.md`](data/raw/README.md) carries the full download table for
these and for the other data used in the analysis — the territorial rule
database, OWID population and the Maddison Project Database — with the version
used, download date, licence and link per file.

Most raw files are not committed: the set is ~200 MB and PRIMAP-hist alone
exceeds GitHub's 100 MB per-file limit. `scripts/00_download_raw.py` fetches
every source that offers a stable direct link. The three that do not — Climate
Watch, which has no direct link, and the two Global Carbon Project
spreadsheets, which sit behind a licence-accept click-through — are committed
in `data/raw/`, so **a fresh clone runs with no manual downloads**. All three
are CC BY 4.0, which permits redistribution with attribution.

Two sources update their published files in place (OWID, Climate Watch) so
a fresh download can differ from the vintage this analysis uses, and ICOS now
serves v1.0 of the Global Carbon Project spreadsheets where this analysis used
v0.3 and v0.2. The committed outputs were generated from the versions recorded
in `data/raw/README.md`.

**The complete set of raw data files, at the exact versions used in the
analysis, is available for download from the [working paper's page on
cgdev.org](https://www.cgdev.org/publication/comparative-analysis-greenhouse-gas-emissions-developed-and-developing-countries-1850).**
Use that in preference to `00_download_raw.py` if you want to reproduce the
committed outputs exactly rather than re-run the analysis against current data.

## Measures and source coverage

All sources are standardised to the schema
`source, iso_code, year, measure, value_Mt`, where `measure` is one of:

| Code | Meaning |
|------|---------|
| `co2_excLUC` | CO2, excluding land-use change |
| `co2_incLUC` | CO2, including land-use change |
| `ghg_excLUC` | All GHGs (AR5 GWP-100), excluding land-use change |
| `ghg_incLUC` | All GHGs (AR5 GWP-100), including land-use change |
| `gmst_*` | Temperature response (°C) matched to each of the four measures — GMST source only, not comparable to emissions |

| Source | co2_excLUC | co2_incLUC | ghg_excLUC | ghg_incLUC | gmst_* |
|--------|-----------|-----------|-----------|-----------|--------|
| OWID | 1750–2024 | 1750–2024 | 1750–2024 | 1750–2024 | — |
| ClimateWatch | 1990–2023 | 1990–2023 | 1990–2023 | 1990–2023 | — |
| PRIMAP | 1750–2024 | 1750–2024 | 1750–2024 | 1750–2024 | — |
| EDGAR | 1970–2024 | 1990–2024 | 1970–2024 | 1990–2024 | — |
| GCP_fossil | 1850–2024 | — | — | — | — |
| GCP_BLUE | — | 1850–2024 | — | — | — |
| GCP_OSCAR | — | 1850–2024 | — | — | — |
| GCP_LUCE | — | 1850–2024 | — | — | — |
| GMST | — | — | — | — | 1851–2024 |

**OWID's CO2 including land-use change is derived here, not taken as
published.** OWID's own `co2_including_luc` column is NaN whenever fossil
`co2` is NaN, even where the land-use estimate exists. The pipeline computes
`co2_incLUC = co2 (fill 0) + land_use_change_co2 (fill 0)`, NaN only when both
are missing, which recovers 13,539 country-year rows across 181 countries —
largest by volume: Indonesia (~9.1 Gt), Myanmar (~8.1 Gt), China (~7.1 Gt).
Our OWID `co2_incLUC` figures therefore differ from OWID's published column by
design.

**World totals here are the sum of countries, not any source's "World" row.**
Every share in this analysis uses `annex_1 + non_annex_1` as the denominator.
Sources define their own World rows differently, and most include
international aviation and shipping, which none of them allocate to individual
countries. Summing countries keeps the denominator consistent across sources;
the trade-off is that shares are of country-attributable emissions only, with
bunker fuels outside the total.

**Colonial attribution** (Figure 8, and the interactive's toggle) reassigns
each country-year's emissions to whoever controlled the territory that year,
using Carbon Brief's (territory, year, power) coefficients. The coefficients
sum to 1 per territory-year, so world totals are unchanged; territories absent
from the database keep their own emissions. GMST is reattributed by
differencing each cumulative series into annual increments, reassigning those,
and re-cumulating — a linear approximation that behaves well for cumulative
1850–2024 comparisons but should not be read year by year.

## Interactive chart

An interactive dumbbell chart comparing each source's estimate of cumulative
emission shares (Annex I / Annex II / non-Annex I), with a measure selector,
colonial-attribution toggle, year-range slider, and a custom country-group
picker is at `ghg-emissions-interactive/`.

Embed URL, configuration parameters, colours, payload sizes and the toolkit
delivery checklist are documented in
[`ghg-emissions-interactive/README.md`](ghg-emissions-interactive/README.md).

## Repository map

```
ghg-emissions/
  README.md                  # this file
  LICENSE                    # MIT — governs except where a subdirectory overrides
  TRACKING.md                # analytics event inventory
  main.py                    # pipeline runner
  requirements.txt
  data-registry.md           # inventory of every input and output
  .nojekyll                  # serve GitHub Pages without a Jekyll pass
  data/
    country_groups.csv       # curated group membership
    raw/                     # raw sources (download links in its README)
      LICENSE                # third-party notice — CGD is not the licensor
    outputs/                 # CC BY-NC-SA 4.0 (see its LICENSE)
      Annex_C.xlsx           # summary tables
      charts/                # chart-input CSVs + Figure_1..11.xlsx
  scripts/
    00_download_raw.py       # fetch raw sources
    00_extract_country_groups.py   # provenance only — needs unpublished workbook
    01..06_*.py              # one extract per source
    07_extract_colonial.py
    08_stack_and_aggregate.py
    09_prepare_chart_data.py
    10_build_summary_tables.py
    11_validate.py
    12_prepare_interactive_dumbbell.py
    charts/                  # PowerShell + Excel COM figure scripts
  templates/                 # vendored CGD chart templates (see its README)
  ghg-emissions-interactive/ # the interactive — CC BY-NC-SA 4.0 (see its LICENSE)
```

## How to reproduce

Requires Python 3 (developed on 3.14) and the packages in `requirements.txt`. The figure
workbooks additionally need Windows with Excel (they are committed, so nobody
has to run that stage).

```bash
pip install -r requirements.txt

# Fetch raw data (no manual steps — the three sources that cannot be
# fetched automatically are committed; see data/raw/README.md)
python scripts/00_download_raw.py

# Run the pipeline: extracts -> colonial coefficients -> stack/aggregate ->
# chart data -> Annex C -> validation -> interactive payloads
python main.py
```

`main.py` flags:

| Flag | Effect |
|------|--------|
| `--download` | run the download script first |
| `--charts` | also rebuild the Excel figures (Windows + Excel, 11 COM scripts) |
| `--only 9,10` | run only the listed step numbers |
| `--from 8` | resume from a step |
| `--skip-validate` | skip the validation step |

Step numbers match the script numbers in `scripts/`. Before step 1, `main.py`
verifies every expected raw file exists and lists anything missing.

Note on regeneration: the committed outputs were produced from the raw-file
vintages recorded in `data/raw/README.md`. OWID and Climate Watch update
their files in place, so regenerating from a fresh download will produce
slightly different numbers.

`scripts/00_extract_country_groups.py` is provenance-only: it documents how
the tracked `data/country_groups.csv` was originally built from an internal
workbook that is not published, and cannot run from a clean clone.

## Outputs

**Committed**:

- `data/outputs/charts/*.csv` — chart-input data for every figure
- `data/outputs/charts/Figure_1.xlsx` … `Figure_11.xlsx` — the figure workbooks
- `data/outputs/Annex_C.xlsx` — the paper's summary tables
- `ghg-emissions-interactive/data.js`, `data.json`, `countries_*.js` — the
  interactive's payloads (GitHub Pages serves them from the tree)

**Regenerated by `main.py`** (gitignored):

- the per-source `*_long.csv` extracts
- `colonial_attribution_long.csv`
- `countries_annual.csv`
- `groups_summary.csv`
- `all_sources_stacked.csv`

## Licence

This repository carries three licences, because it holds three different kinds
of thing.

| Path | Licence | What it covers |
|------|---------|----------------|
| Repository root and everything not listed below | MIT | The pipeline scripts, `main.py`, the chart templates and the documentation — CGD's own work |
| `ghg-emissions-interactive/` | CC BY-NC-SA 4.0 | The interactive and its data payloads |
| `data/outputs/` | CC BY-NC-SA 4.0 | Chart data, figure workbooks and the Annex C tables |
| `data/raw/` | Third-party, CC BY 4.0 | Three committed source files that are not CGD's work — see `data/raw/LICENSE` |

**Precedence:** the MIT licence at the repository root governs, except where a
subdirectory contains its own `LICENSE` file, which governs that directory and
everything beneath it. GitHub's licence detection only reads the root, so the
repository is labelled MIT in the sidebar; the table above is the full picture.

**Why the outputs are more restrictive than the code.** They incorporate values
derived from PRIMAP-hist, which is licensed CC BY-NC-SA 4.0. Its ShareAlike
term requires adapted material to be released under the same licence, and its
NonCommercial term carries across with it. The scripts contain no source data,
so that obligation does not reach them — anyone is free to take the MIT-licensed
pipeline, download the sources themselves and generate their own outputs.

Attribution for every source, with versions, download dates and links, is in
[`data/raw/README.md`](data/raw/README.md).

## Contact

Edward Wickstead (ewickstead@cgdev.org) —
[edward-wickstead-cgd](https://github.com/edward-wickstead-cgd)
