# Raw source data

Raw source files are **not committed** (the largest exceeds GitHub's 100 MB per-file
limit; the full set is ~200 MB), with one exception noted below. Download them into
this folder using the links here, or run `python scripts/00_download_raw.py` from the
repository root, which fetches every source that offers a direct link and prints
instructions for the rest.

All links were verified on 11 August 2026. Two of the sources (OWID, Climate Watch)
update their published files in place, so a fresh download may differ from the
version used in this analysis — the version/date columns below record exactly what
we used. The committed outputs in `data/outputs/` were generated from these versions.

| File | Source | Version used | Downloaded | Licence | Link |
|------|--------|--------------|------------|---------|------|
| `owid-co2-data.csv` | Our World in Data | GCB 2025 vintage | 4 Feb 2026 | CC BY 4.0 | <https://owid-public.owid.io/data/co2/owid-co2-data.csv> (also in the [owid/co2-data](https://github.com/owid/co2-data) repo) |
| `owid-co2-codebook.csv` | Our World in Data | — | 4 Feb 2026 | CC BY 4.0 | <https://owid-public.owid.io/data/co2/owid-co2-codebook.csv> |
| `CW_HistoricalEmissions_ClimateWatch.csv` | Climate Watch (WRI) | 1990–2023 release | 20 Apr 2026 | CC BY 4.0 | **Committed in this folder** — no stable direct link exists (export via the [GHG emissions explorer](https://www.climatewatchdata.org/ghg-emissions)) |
| `Guetschow_et_al_2025a-PRIMAP-hist_v2.7_final_no_extrap_no_rounding_22-Aug-2025.csv` | PRIMAP-hist (Gütschow et al.) | v2.7 (22 Aug 2025 build) | 24 Nov 2025 | CC BY-NC-SA 4.0 | Zenodo DOI [10.5281/zenodo.17090760](https://doi.org/10.5281/zenodo.17090760) |
| `EDGAR_2025_GHG_booklet_2025.xlsx` | EDGAR Community GHG Database (JRC/IEA), 2025 report | 2025 edition (AR5 GWPs) | 1 May 2026 | CC BY 4.0 | <https://edgar.jrc.ec.europa.eu/booklet/EDGAR_2025_GHG_booklet_2025.xlsx> |
| `National_Fossil_Carbon_Emissions_2025_v0.3.xlsx` | Global Carbon Project (GCB 2025) | v0.3 | 19 Feb 2026 | CC BY 4.0 | [ICOS data supplement](https://www.icos-cp.eu/impact/science/global-carbon-budget/2025) — licence-accept click-through; ICOS has since released v1.0 |
| `National_LandUseChange_Carbon_Emissions_2025v0.2.xlsx` | Global Carbon Project (GCB 2025) | v0.2 | 19 Feb 2026 | CC BY 4.0 | [ICOS data supplement](https://www.icos-cp.eu/impact/science/global-carbon-budget/2025) — licence-accept click-through; ICOS has since released v1.0 |
| `GMST_response_1851-2024.csv` | Jones et al., National contributions to climate change | v2025.1 | 24 Feb 2026 | CC BY 4.0 | Zenodo DOI [10.5281/zenodo.16640595](https://doi.org/10.5281/zenodo.16640595) |
| `territorial_rule_database_1850_2023.csv` | Carbon Brief, colonial-emissions-data | 1850–2023 | 27 Apr 2026 | none stated | <https://raw.githubusercontent.com/carbonbrief/colonial-emissions-data/main/output-clean/territorial_rule_database_1850_2023.csv> |
| `population.csv` | Our World in Data (HYDE; Gapminder; UN WPP) | July 2024 update | 19 May 2026 | CC BY 4.0 | Grapher CSV export: <https://ourworldindata.org/grapher/population.csv?v=1&csvType=full&useColumnShortNames=false> |
| `population.metadata.json` | Our World in Data | — | 19 May 2026 | CC BY 4.0 | Grapher metadata export (same page as above) |
| `mpd2023_web.xlsx` | Maddison Project Database 2023 (Bolt & van Zanden) | 2023 release | 20 May 2026 | CC BY 4.0 | <https://dataverse.nl/api/access/datafile/421302> ([release page](https://www.rug.nl/ggdc/historicaldevelopment/maddison/releases/maddison-project-database-2023)) |

Notes:

- **Climate Watch** is committed as the single exception because its explorer export
  has no stable URL. Its CC BY 4.0 licence permits redistribution with attribution.
- **PRIMAP-hist** is CC BY-NC-SA 4.0 (non-commercial): download it from the Zenodo
  record; it is not redistributed here. The copy used in this analysis stores very
  small values in fixed-decimal notation (it appears to come from the PRIMAP-hist
  website rather than Zenodo); compared cell-by-cell against the Zenodo file on
  11 Aug 2026, the maximum difference is 5e-8 Gg — negligible at the Mt scale of
  every output.
- **GCP national emissions**: we used v0.3 (fossil) and v0.2 (land-use change),
  accessed 19 February 2026. ICOS now serves v1.0 of both; numbers may differ
  slightly from the committed outputs if you regenerate with v1.0.
- **Carbon Brief territorial rule database** states no licence; it is linked, not
  redistributed. It supplies the (territory, year, power) coefficients used to
  reassign country-year emissions to the powers that controlled them — Figure 8 and
  the interactive's colonial-attribution toggle.
- `population.csv` / `mpd2023_web.xlsx` are used only to patch Ireland's population
  before 1950, which OWID does not carry: the "Ireland (whole island)" series for
  1850–1920 and the Maddison Project Database for 1921–1949.
