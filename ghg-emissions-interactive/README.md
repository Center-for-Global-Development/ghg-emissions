# The interactive

An interactive dumbbell chart comparing each source's estimate of cumulative
emission shares (Annex I / Annex II / non-Annex I), with a measure selector,
colonial-attribution toggle, year-range slider, and a custom country-group
picker.

**Embedded on:**
<https://www.cgdev.org/blog/not-fit-purpose-why-climate-classifications-developed-and-developing-countries-have-change>

**Embed URL:**
`https://center-for-global-development.github.io/ghg-emissions/ghg-emissions-interactive/`

**Recommended iframe** (paste into the CMS; the page's resize listener sets the
height after load, so the `height` here is only a pre-load placeholder):

```html
<iframe
  src="https://center-for-global-development.github.io/ghg-emissions/ghg-emissions-interactive/"
  title="Interactive chart: Annex I and Annex II shares of world cumulative emissions, by dataset"
  width="100%"
  height="720"
  style="border: 0; width: 100%; max-width: 100%; display: block;"
  loading="lazy"
  scrolling="no"></iframe>
```

Append URL parameters to `src` to change the initial view — for example
`?custom=0` hides the country-group picker for a blog embed, and
`?measure=co2_incLUC&start=1990&end=2024` opens on a different measure and
period. See "Embed configuration" below for the full list.

> The interactive is in this subfolder, not at the repository root. The CGD
> Interactive Toolkit's single-interactive layout puts `index.html` at the
> root; this repository also holds the full analysis pipeline, so the
> interactive sits in a subfolder. All asset paths are relative and the page
> runs unchanged from a subpath. CGD comms have confirmed this layout works
> with the embed process (August 2026).

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
come from `data/country_groups.csv` via script 12: annex, institutional
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
entry in script 12's `MEMBERSHIP_SETS` — the picker builds its modes from
`data.setGroups`.

**Analytics.** `tracking.js` is the toolkit standard's utility with slug
`ghg-emissions-interactive`; every event is inventoried in `TRACKING.md` at
the repository root. `PARENT_ORIGIN` is `https://www.cgdev.org`.

**Fonts.** The page loads Sofia Pro through CGD's Adobe Fonts kit
(`use.typekit.net/ymp6ujv.css`, provided by CGD comms) and falls back to a
locally installed Sofia Pro, then Segoe UI / system sans — the fallback chain
the CGD brand reference sanctions. The kit is the page's only external
request; if it is unreachable the page still renders on the fallback fonts.

## Payload sizes

`data.js` is ~203 KB (group series + picker metadata). The four
`countries_*.js` payloads are 1.6–2.9 MB each (~0.5–1 MB gzipped) and carry
per-country annual series for every source under both attributions — they
cannot be meaningfully reduced without dropping picker functionality. They are
lazy-loaded on first use of the country picker, so an embed that never opens
it (or sets `?custom=0`) never fetches them.

Both payloads are built by `scripts/12_prepare_interactive_dumbbell.py` and
are committed — GitHub Pages serves them straight from the tree.

## Toolkit delivery checklist

| Check | Status |
|---|---|
| Iframe resize script present, `PARENT_ORIGIN` not `'*'` | Done |
| Analytics implemented, `TRACKING.md` current | Done |
| No outer chrome; transparent background; no fixed height | Done |
| Responsive 320–1200 px without horizontal scroll | Done |
| Palette from `cgd-brand-reference.md` | Done |
| Keyboard operable, visible focus states, chart-equivalent data table | Done |
| No build step, no CDN dependencies, no `innerHTML` | Done — the one external request is the comms-provided Adobe Fonts kit |
| Data sources, dates and transformations documented | Done — the root `README.md` + `data-registry.md` |
| Static title/subtitle/source note moved to CMS text | Kept in the page — parts of the text change dynamically with the chart settings, and no CMS copy exists; noted as a deliberate deviation |
| Own repo under the CGD org | Done — this repository |
