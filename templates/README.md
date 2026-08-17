# Chart templates

These four Excel templates are vendored copies of the CGD chart templates
maintained in the internal `cgd-general` repository
(`edward-wickstead-cgd/cgd-general`, `chart-templates/`), copied on
11 August 2026. They carry CGD branding (colours, Sofia Pro/Calibri typography,
axis and legend styling) and are consumed by the PowerShell chart scripts in
`scripts/charts/`, which copy a template, inject data, and re-wire the series.

| Template | Used by |
|----------|---------|
| `Chart-template-Dumbbell.xlsx` | fig3, fig4, fig8, fig10, fig11 |
| `Chart-template-MultiLine.xlsx` | fig6, fig7 |
| `Chart-template-PiePair.xlsx` | fig1 |
| `Chart-template-StackedArea.xlsx` | fig2, fig5, fig9 |

If the upstream templates in `cgd-general` are fixed or improved, re-copy them
here — these files do not update automatically.
