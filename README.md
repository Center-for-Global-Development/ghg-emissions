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

## Interactive chart

An interactive dumbbell chart comparing each source's estimate of cumulative
emission shares (Annex I / Annex II / non-Annex I), with a measure selector,
colonial-attribution toggle, year-range slider, and a custom country-group
picker is at `ghg-emissions-interactive/`.

Embed URL, configuration parameters, colours, payload sizes and the toolkit
delivery checklist are documented in
[`ghg-emissions-interactive/README.md`](ghg-emissions-interactive/README.md).

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
used, download date, licence and link per file. Raw files are not committed
(the set is ~200 MB and PRIMAP-hist alone exceeds GitHub's 100 MB per-file
limit); `scripts/00_download_raw.py` fetches everything that offers a direct
link.

Two sources update their published files in place (OWID, Climate Watch) so
a fresh download can differ from the vintage this analysis uses. The committed
outputs were generated from the versions recorded in `data/raw/README.md`.

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

## How to reproduce

Requires Python 3 (developed on 3.14) and the packages in `requirements.txt`. The figure
workbooks additionally need Windows with Excel (they are committed, so nobody
has to run that stage).

```bash
pip install -r requirements.txt

# Fetch raw data (GCP requires one manual licence-accept download — the
# script prints instructions; see data/raw/README.md)
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

## Repository map

```
ghg-emissions/
  README.md                  # this file
  TRACKING.md                # analytics event inventory
  main.py                    # pipeline runner
  requirements.txt
  data-registry.md           # inventory of every input and output
  .nojekyll                  # serve GitHub Pages without a Jekyll pass
  data/
    country_groups.csv       # curated group membership
    raw/                     # raw sources (download links in its README)
    outputs/
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
  ghg-emissions-interactive/ # the interactive (page code + data payloads)
```

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
- `Summary_tables_v6.xlsx` (the working copy of the Annex C tables)

## Contact

Edward Wickstead, Center for Global Development —
[edward-wickstead-cgd](https://github.com/edward-wickstead-cgd)
