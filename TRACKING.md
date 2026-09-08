# Event Tracking: Cumulative emissions shares (interactive dumbbell)

`interactive_name`: `ghg-emissions-interactive`

Tracking implemented per the [CGD Interactive Analytics Tracking Standard](https://github.com/Center-for-Global-Development/cgd-interactive-toolkit/blob/main/analytics-tracking-standard.md).

The slug is set as the default in `tracking.js`. This project currently contains a
single interactive, so no `window.CGD_INTERACTIVE_NAME` is set in `index.html`. If
further figures are added as separate iframes, each HTML file must set that global
before loading `tracking.js` and this file must gain a table per interactive.

## Tracked Events

`interactive_view` fires once, after the first render.

| `action_type` | `action_label` | `action_value` | Notes |
|---|---|---|---|
| `filter` | `measure_select` | `ghg_incLUC`, `co2_incLUC`, `ghg_excLUC`, `co2_excLUC` | Measure dropdown. 4 values. |
| `filter` | `colonial_attribution` | `on`, `off` | Colonial attribution toggle. 2 values. |
| `filter` | `year_range` | decade-rounded range, e.g. `1900–2000`; `1850–2024` when at the full range | Dual-handle year slider. Fires on `change` (pointer release), not on the `input` events that drive rendering, and is debounced (800 ms) so a keyboard user stepping through years sends one event, not one per key press. Start and end are floored to the decade unless the handle sits at the slider bound, where the true bound is kept. Both handles share one label. At most ~180 distinct values. |
| `preset` | `year_range_reset` | — | Reset button. Disabled at full range, so it only fires on a real reset. |
| `detail_open` / `detail_close` | `country_picker` | — | "+ Add country group" opens; Cancel closes. Adding a group closes the picker without a `detail_close`. |
| `view_control` | `picker_group_by` | `none`, `annex`, `institutional`, `development`, `income`, `region` | Group-by dropdown inside the picker. 6 values. |
| `filter` | `add_country_group` | named set only, e.g. `G7`, `EU27`, `BRICS` | Fires on "Add to chart". `action_value` is sent **only** when the selection exactly matches a known named set; user-typed and default (`Custom group N`) names are omitted, per the cardinality rule. |
| `filter` | `remove_country_group` | — | Fires from both removal paths: the × on the group chip and the "remove ×" control on the chart row. Group names are free text so no value is sent. |
| `detail_open` / `detail_close` | `country_group_members` | — | Expanding a group chip to list its member countries. |
| `detail_open` / `detail_close` | `data_table` | — | "View the data" `<details>` element. |

## Not Tracked

- **Tooltip hover on chart marks.** Excluded by the standard — high volume, low signal.
  The marks are also focusable, so tracking would fire on keyboard traversal too.
- **Slider `input` events during drag.** Rendering listens on `input`; tracking listens
  on `change` only and debounces, so one adjustment produces one event.
- **Exact slider years.** The tracked `year_range` is decade-rounded (see above) to keep
  the value set bounded.
- **Picker search box.** Free text, unbounded cardinality.
- **Per-country checkboxes and bucket select-all inside the picker.** High volume during
  a single selection; the resulting `add_country_group` event captures the outcome.
- **Group chip name field.** Free text, and captured indirectly by `add_country_group`.

## Maintenance

`app.js` routes every event through a local `track()` wrapper (defined just above
`svgEl`), which no-ops if `tracking.js` failed to load. Any PR that adds, removes or
renames a tracked interaction must update this file in the same PR.
