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

A link to the published paper will be added here when it is live.

## The interactive

An interactive dumbbell chart comparing each source's estimate of cumulative
emission shares (Annex I / Annex II / non-Annex I), with a measure selector,
colonial-attribution toggle, year-range slider, and a custom country-group
picker.

Embed URL:
`https://center-for-global-development.github.io/ghg-emissions/ghg-emissions-interactive/`

> The interactive is at `ghg-emissions-interactive/`, not at the repository
> root. The CGD Interactive Toolkit's single-interactive layout puts
> `index.html` at the root; this repository also holds the full analysis
> pipeline, so the interactive sits in a subfolder. All asset paths are
> relative and the page runs unchanged from a subpath. CGD comms have
> confirmed this layout works with the embed process (August 2026).

Full documentation of the interactive is in the
[Interactive documentation](#interactive-documentation) section below.

## Data sources

| Dataset | Source | Coverage |
|---------|--------|----------|
| OWID CO2 data | Our World in Data (Global Carbon Budget 2025) | 1750–2024, CO2 + all GHGs |
| Climate Watch | Climate Watch Historical GHG Emissions (WRI) | 1990–2023, all GHGs |
| PRIMAP-hist v2.7 | Gütschow et al. (2025) | 1750–2024, all GHGs |
| EDGAR 2025 booklet | JRC / IEA EDGAR Community GHG Database | 1970–2024, CO2 + all GHGs excl. LULUCF; 1990–2024 incl. LULUCF |
| GCP (fossil) | Global Carbon Project (GCB 2025) | 1850–2024, fossil CO2 |
| GCP (land-use change) | Global Carbon Project (GCB 2025) | 1850–2024, LUC CO2 (3 models: BLUE, OSCAR, LUCE) |
| GMST | Jones et al., National contributions to climate change (v2025.1) | 1851–2024, temperature response (°C) |
| Territorial rule database | Carbon Brief, colonial-emissions-data | 1850–2023, colonial-rule coefficients |

**Raw source files are not committed** (with one exception). The largest
(PRIMAP-hist v2.7, ~112 MB) exceeds GitHub's 100 MB per-file limit and the
full set is ~200 MB. [`data/raw/README.md`](data/raw/README.md) carries the
full download table — link, version used, download date, and licence per
source — and `scripts/00_download_raw.py` fetches everything that offers a
direct link. The exception is the Climate Watch extract, which is committed
because its explorer export has no stable URL (its CC BY 4.0 licence permits
redistribution with attribution).

Two sources update their published files in place (OWID, Climate Watch), and
ICOS has released newer versions of the GCP files than the ones used here, so
a fresh download can differ from the vintage this analysis used. The committed
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

## Method

**Country groups.** Countries carry binary flags for 25 groups (Annex I,
Annex II, non-Annex I, income groups, World Bank regions, G7/G20/EU/OECD/
BRICS, LDC/LLDC/SIDS, …) in the tracked `data/country_groups.csv`. Group
totals sum member countries.

**World denominator.** Throughout the pipeline the world total is
`annex_1 + non_annex_1` — the sum of individual countries — rather than any
source's own "World" row. This keeps shares consistent across sources and
excludes international aviation and shipping, which no source allocates to
countries (see caveats).

**Cumulative per-capita emissions** (Figures 10/11) are computed three ways,
following the framing in Beynon (2024):

- **Trajectory-summed (TS)** — the sum across years of each year's per-capita
  emissions: `Σ_y E_y / P_y`. Each year's rate uses that year's population.
- **Endpoint-normalised (EN)** — cumulative emissions divided by end-year
  population: `Σ_y E_y / P_end`.
- **Person-year-weighted (PYW)** — cumulative emissions divided by cumulative
  person-years: `Σ_y E_y / Σ_y P_y`. The only rule that aggregates
  self-consistently across both countries and years.

Group figures apply each formula to group totals (sum of numerators over sum
of denominators), not to averages of country values.

**Colonial attribution** (Figure 8 and the interactive's toggle). Country-year
emissions are reassigned to the controlling powers using Carbon Brief's
territorial-rule coefficients: emissions are multiplied by each
(territory, year, power) coefficient and re-aggregated on the attributed-to
country. Coefficients sum to 1 per territory-year, so world totals are
preserved exactly (verified to floating-point precision across all
source × measure combinations). Countries absent from the attribution table
remain their own. The coefficient series is forward-filled from 2023 to 2024.

GMST is reattributed by a linear approximation: each country's cumulative
temperature-response series is differenced into annual increments, each
increment is reassigned using that year's coefficient, and the result is
re-cumulated. Each increment mixes that year's new warming with decay of prior
warming, so this does not exactly reproduce a re-run of the underlying climate
response model; for cumulative 1850–2024 comparisons it behaves well, but
individual transition years should not be over-interpreted.

**GMST aggregation.** Group GMST values are aggregated from Jones et al.'s
country rows, not their pre-aggregated group rows, whose Annex I definition
excludes Kazakhstan and Monaco. We use the current UNFCCC Annex I list
(including both), consistent with every other source in the pipeline; the
difference is ~0.014 °C (Kazakhstan's cumulative contribution).

## Data handling and caveats

- **Aggregate rows dropped.** Climate Watch's `EUU` (EU bloc) row would
  double-count EU members and is excluded. The GMST source's group codes
  (`EIT`, `OECD`, `EU27`, `EU28`, `AFRICA`, `ASIA`, …) are likewise excluded
  from country-level data. `scripts/12_validate.py` checks that no known
  aggregate code leaks into the country-level output.
- **OWID CO2 incl. LUC is derived, not taken as published.** OWID's
  `co2_including_luc` column is NaN whenever fossil `co2` is NaN, even if the
  land-use estimate exists. The pipeline computes
  `co2_incLUC = co2 (fill 0) + land_use_change_co2 (fill 0)` (NaN only when
  both are missing), recovering 13,539 country-year rows across 181 countries
  — largest: Indonesia (~9.1 Gt), Myanmar (~8.1 Gt), China (~7.1 Gt).
- **GCP fossil × LUC merge is an outer join** with missing values filled to
  zero. An inner join would silently drop fossil-only territories (Taiwan,
  Hong Kong, 18 smaller ones) and — far larger — the ~135 countries that have
  LUC but no fossil entry in the earliest years.
- **International aviation and shipping** are excluded from country totals in
  every source (separate rows / bunker-fuel memo items / AIR & SEA entities),
  and therefore from all group totals and the world denominator here.
- **ANT (Netherlands Antilles)** has PRIMAP/EDGAR data but no entry in
  `country_groups.csv`. The as-reported group totals include it in
  non-Annex I; the colonial-attributed group totals drop it (~0.08% of PRIMAP
  non-Annex I) — a known, minor inconsistency. The interactive's country
  picker includes it explicitly.
- **GMST-only ISO codes** (`KSV`, `XKW`, `XPC`, `XRY`) are excluded from both
  group totals and the interactive's picker.
- **Ireland's population before 1950** is missing in OWID. Per-capita figures
  patch it from OWID's "Ireland (whole island)" series for 1850–1920 (minor
  overcount: includes Northern Ireland) and the Maddison Project Database 2023
  for 1921–1949. OWID's GBR series was checked for a partition-era step change
  and tracks Great Britain + Northern Ireland throughout, so there is no
  double-counting.
- **PRIMAP LULUCF has a splice discontinuity at 1989→1990** (~5 Gt drop) where
  the underlying source changes; it propagates to `ghg_incLUC` comparisons
  across that boundary.
- **Final-year coverage is partial.** Climate Watch ends in 2023, and
  PRIMAP-hist's 2024 estimates cover only a subset of countries. Annual-share
  figures can show an end-year artefact; end years should not be compared
  directly with earlier years.
- **Colonial attribution upstream quirk.** Carbon Brief's database attributes
  South Sudan 100% to the UK from 1882 through 2023 with no transition to
  independence (unlike Sudan, which transitions in 1956). Retained as
  published; interpret UK-attributed shares accordingly.
- **Per-capita NaN rule (Figure 9).** Group population denominators include
  every member with population data, whether or not it has emissions data
  (a pre-industrial country still has people); numerators include only
  countries with both. This prevents artificial jumps when a large country
  enters a measure's coverage.

## How to reproduce

Requires Python 3 (developed on 3.14) and the packages in `requirements.txt`. The figure
workbooks additionally need Windows with Excel (they are committed, so nobody
has to run that stage).

```bash
pip install -r requirements.txt

# Fetch raw data (GCP requires one manual licence-accept download — the
# script prints instructions; see data/raw/README.md)
python scripts/00_download_raw.py

# Run the pipeline: extracts -> stack/aggregate -> workbook -> chart data ->
# Annex C -> validation -> interactive payloads
python main.py
```

`main.py` flags:

| Flag | Effect |
|------|--------|
| `--download` | run the download script first |
| `--charts` | also rebuild the Excel figures (Windows + Excel, ~14 COM scripts) |
| `--only 8,10` | run only the listed step numbers |
| `--from 8` | resume from a step |
| `--skip-validate` | skip the validation step |

Before step 1, `main.py` verifies every expected raw file exists and lists
anything missing.

Note on regeneration: the committed outputs were produced from the raw-file
vintages recorded in `data/raw/README.md`. OWID and Climate Watch update
their files in place, so regenerating from a fresh download will produce
slightly different numbers — that is source revision, not pipeline
non-determinism.

`scripts/00_extract_country_groups.py` is provenance-only: it documents how
the tracked `data/country_groups.csv` was originally built from an internal
workbook that is not published, and cannot run from a clean clone.

## Repository map

```
ghg-emissions/
  README.md                  # this file
  TRACKING.md                # analytics event inventory (toolkit requirement)
  main.py                    # pipeline runner
  requirements.txt
  data-registry.md           # inventory of every input and output
  .nojekyll                  # serve GitHub Pages without a Jekyll pass
  data/
    country_groups.csv       # curated group membership (committed)
    raw/                     # raw sources (download links in its README;
                             #   only the Climate Watch extract is committed)
    outputs/
      Annex_C.xlsx           # summary tables (committed)
      charts/                # chart-input CSVs + Figure_1..13.xlsx,
                             #   Figure_A_dynamic.xlsx (committed)
  scripts/
    00_download_raw.py       # fetch raw sources
    00_extract_country_groups.py   # provenance only — needs unpublished workbook
    01..06_*.py              # one extract per source
    07_stack_and_aggregate.py
    08_build_workbook.py
    09_prepare_chart_data.py
    11_extract_colonial.py
    12_validate.py
    13_build_summary_tables.py
    14_prepare_interactive_dumbbell.py
    charts/                  # PowerShell + Excel COM figure scripts
  templates/                 # vendored CGD chart templates (see its README)
  ghg-emissions-interactive/ # the interactive (page code + data payloads)
```

## Outputs

**Committed** (so a reader never has to run anything):

- `data/outputs/charts/*.csv` — chart-input data for every figure
- `data/outputs/charts/Figure_1.xlsx` … `Figure_13.xlsx`,
  `Figure_A_dynamic.xlsx` — the figure workbooks
- `data/outputs/Annex_C.xlsx` — the paper's summary tables
- `ghg-emissions-interactive/data.js`, `data.json`, `countries_*.js` — the
  interactive's payloads (GitHub Pages serves them from the tree)

**Regenerated by `main.py`** (gitignored): the per-source `*_long.csv`
extracts, `countries_annual.csv`, `groups_summary.csv`,
`all_sources_stacked.csv`, and `Summary_tool_v2.xlsx` — a parameterised
Excel comparison workbook.

## Interactive documentation

Plain HTML + CSS + vanilla JS; no build step, no external dependencies.

**Colours** (from the toolkit's `cgd-brand-reference.md`): OWID `#006970`,
PRIMAP `#FFB52C`, GCP `#2D99B5`, EDGAR `#D15553`, Climate Watch `#00896C`,
GMST marker `#1A272A`. The categorical palette's three pale entries are
unreadable as 12 px dots on white, leaving five hues for five sources; the two
status colours fill the gap.

**Controls.** Measure dropdown (inclusive-LULUCF measures first; default
all-GHG incl. LULUCF, which OWID and PRIMAP cover back to 1850), a
colonial-attribution toggle, and a dual-handle year-range slider (floored at
1850) with Reset. Datasets show automatically when they cover the selected
start year and reach at least (end year − 2) — the tolerance covers Climate
Watch ending in 2023, matching the static Figure 4.

**Custom country groups.** "+ Add country group" opens a picker (search,
group-by, per-group select-all, name field; up to 4 rows). Group-by modes
come from `data/country_groups.csv` via script 14: annex, institutional
(G20/G7/EU27/OECD/BRICS), development (LDC/LLDC/SIDS), income, World Bank
region. Institutional and development sets overlap, so a country appears
under every set it belongs to. Selecting a whole bucket offers its name as
the row label. Custom shares use the same world denominator as the preset
rows; colonial attribution uses country-level precomputed series; GMST sums
the country-level Jones et al. series. Hide the picker in embeds with
`?custom=0` or `EMBED_DEFAULTS.custom = false`.

**Share and GMST maths** (mirrors the Excel figures; QA'd against static
Figures 3/4/8 — emission shares match exactly, GMST within 0.03 pp):

- share = Σ group emissions over the range ÷ (Σ Annex I + Σ non-Annex I) × 100
- GMST = (cum[end] − cum[start−1]) ÷ world same × 100, with cum[start−1] = 0
  when start ≤ 1851

**Embed configuration.** Set the initial view per embed without touching code:
URL parameters (`?measure=ghg_incLUC&colonial=1&start=1900&end=2000&custom=0`)
or `window.EMBED_DEFAULTS = {measure, colonial, start, end, custom}` defined
before `app.js` loads.

**Extending.** Preset rows are the `GROUPS` array at the top of `app.js`;
new membership sets need a flag column in `data/country_groups.csv` plus an
entry in script 14's `MEMBERSHIP_SETS` — the picker builds its modes from
`data.setGroups`.

**Analytics.** `tracking.js` is the toolkit standard's utility with slug
`ghg-emissions-interactive`; every event is inventoried in `TRACKING.md` at
the repository root. `PARENT_ORIGIN` is `https://www.cgdev.org`.

### Toolkit delivery checklist

| Check | Status |
|---|---|
| Iframe resize script present, `PARENT_ORIGIN` not `'*'` | Done |
| Analytics implemented, `TRACKING.md` current | Done |
| No outer chrome; transparent background; no fixed height | Done |
| Responsive 320–1200 px without horizontal scroll | Done |
| Palette from `cgd-brand-reference.md` | Done |
| Keyboard operable, visible focus states, chart-equivalent data table | Done |
| No build step, no CDN dependencies, no `innerHTML` | Done — the one external request is the comms-provided Adobe Fonts kit |
| Data sources, dates and transformations documented | Done — this README + `data-registry.md` |
| Static title/subtitle/source note moved to CMS text | Kept in the page — parts of the text change dynamically with the chart settings, and no CMS copy exists; noted as a deliberate deviation |
| Own repo under the CGD org | Done — this repository |

### Payload sizes

`data.js` is ~203 KB (group series + picker metadata). The four
`countries_*.js` payloads are 1.6–2.9 MB each (~0.5–1 MB gzipped) and carry
per-country annual series for every source under both attributions — they
cannot be meaningfully reduced without dropping picker functionality. They are
lazy-loaded on first use of the country picker, so an embed that never opens
it (or sets `?custom=0`) never fetches them.

**Fonts.** The page loads Sofia Pro through CGD's Adobe Fonts kit
(`use.typekit.net/ymp6ujv.css`, provided by CGD comms) and falls back to a
locally installed Sofia Pro, then Segoe UI / system sans — the fallback chain
the CGD brand reference sanctions. The kit is the page's only external
request; if it is unreachable the page still renders on the fallback fonts.

## Licence and citation

No licence file is committed — CGD comms have opted to launch without one
(August 2026), which leaves the repository "all rights reserved" by default.
Each upstream dataset keeps its own licence, listed per source in
[`data/raw/README.md`](data/raw/README.md) — note PRIMAP-hist is
CC BY-NC-SA 4.0 (non-commercial) and the Carbon Brief territorial-rule
database states no licence, which is why neither is redistributed here.

Citation for the analysis will be added when the paper is published.

## Contact

Edward Wickstead, Center for Global Development —
[edward-wickstead-cgd](https://github.com/edward-wickstead-cgd)
