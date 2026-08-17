<#
.SYNOPSIS
figG.ps1
Generate Figure_9.xlsx -- per-capita GHG emissions by country group, 4 tabs.

Inputs:
  data/outputs/charts/fig9_data.csv
  templates/Chart-template-StackedArea.xlsx

Output:
  data/outputs/charts/Figure_9.xlsx

Pre-req: script 09 must have run to produce fig9_data.csv.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

. "$PSScriptRoot\_chart_helpers.ps1"

$G_GROUPS = @("annex_2","annex_1","non_annex_1","world")
$G_LABELS = @{
    "annex_2"     = "Annex II"
    "annex_1"     = "Annex I"
    "non_annex_1" = "Non-Annex I"
    "world"       = "World"
}
$G_COLORS = @{
    "annex_2"     = 5983243   # dark teal  (#0B4C5B) — Annex II
    "annex_1"     = 10390042  # mid teal   (#1A8A9E) — Annex I
    "non_annex_1" = 2930175   # amber      (#FFB52C) — Non-Annex I (developing)
    "world"       = 0         # black
}

function Write-GData {
    param($ws, [string]$measure, $allRows, [string]$title)
    try { $null = $ws.Cells.UnMerge() } catch {}
    $ws.Range("A1:Z300").ClearContents()
    try { $ws.Range("A1:E1").Merge() | Out-Null } catch {}
    $ws.Cells(1,1).Value2 = $title
    try {
        $ws.Cells(1,1).Font.Bold = $true
        try { $ws.Cells(1,1).Font.Name = "Sofia Pro" } catch { $ws.Cells(1,1).Font.Name = "Calibri" }
        $ws.Cells(1,1).Font.Size = 10
    } catch {}
    $ws.Cells(2,1).Value2 = "Year"
    for ($gi = 0; $gi -lt $G_GROUPS.Count; $gi++) {
        $ws.Cells(2, $gi + 2).Value2 = $G_LABELS[$G_GROUPS[$gi]]
    }
    try {
        $ws.Range("A2:E2").Font.Bold = $true
        try { $ws.Range("A2:E2").Font.Name = "Sofia Pro" } catch { $ws.Range("A2:E2").Font.Name = "Calibri" }
        $ws.Range("A2:E2").Font.Size = 9
    } catch {}
    $filtered = @($allRows | Where-Object { $_.measure -eq $measure })
    $r = 3
    foreach ($row in $filtered) {
        $ws.Cells($r, 1).Value2 = [int]$row.year
        for ($gi = 0; $gi -lt $G_GROUPS.Count; $gi++) {
            Set-Value $ws.Cells($r, $gi + 2) $row.($G_GROUPS[$gi])
        }
        $r++
    }
    $endR = $r - 1
    Write-Host "  [$measure] Wrote $($filtered.Count) rows (rows 3-${endR})"
    try {
        $ws.Range("A3:A${endR}").NumberFormat = "0"
        $ws.Range("B3:E${endR}").NumberFormat = "#,##0.000"
        try { $ws.Range("A3:E${endR}").Font.Name = "Sofia Pro" } catch { $ws.Range("A3:E${endR}").Font.Name = "Calibri" }
        $ws.Range("A3:E${endR}").Font.Size = 9
        try { $ws.Rows("$($endR+1):300").Delete() | Out-Null } catch {}
    } catch {}
    return $endR
}

function Set-GChart {
    param($ch, $ws, [string]$yLabel)
    $dn   = $ws.Name
    $endR = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
    try { $ch.ChartType = 4 } catch {}
    # Remove extra series left over from template before re-wiring
    while ($ch.SeriesCollection().Count -gt $G_GROUPS.Count) {
        $ch.SeriesCollection($ch.SeriesCollection().Count).Delete() | Out-Null
    }
    for ($gi = 1; $gi -le $G_GROUPS.Count; $gi++) {
        $col = [char](64 + $gi + 1)
        try { $s = $ch.SeriesCollection($gi) } catch { $s = $ch.SeriesCollection().NewSeries() }
        $s.Name    = "='${dn}'!`$${col}`$2"
        $s.Values  = $ws.Range("${col}3:${col}${endR}")
        $s.XValues = $ws.Range("A3:A${endR}")
        $s.MarkerStyle = -4142
        $color = $G_COLORS[$G_GROUPS[$gi-1]]
        try { $s.Format.Line.ForeColor.RGB = $color } catch {}
        try { $s.Format.Line.Weight = 1.5 } catch {}
    }
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = [string]$ws.Cells(1,1).Value2
    try { $ch.ChartTitle.Font.Name = "Sofia Pro" } catch { try { $ch.ChartTitle.Font.Name = "Calibri" } catch {} }
    try { $ch.ChartTitle.Font.Size = 14 } catch {}
    Set-SubscriptCO2 $ch.ChartTitle
    try {
        $xAx = $ch.Axes(1)
        $xAx.TickLabelSpacing = 10; $xAx.TickMarkSpacing = 10
        try { $xAx.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx.TickLabels.Font.Name = "Calibri" }
        $xAx.TickLabels.Font.Size = 12
    } catch {}
    try {
        $yAx = $ch.Axes(2)
        $yAx.HasTitle = $true
        $yAx.AxisTitle.Text = $yLabel
        try { $yAx.AxisTitle.Font.Name = "Sofia Pro" } catch { $yAx.AxisTitle.Font.Name = "Calibri" }
        $yAx.AxisTitle.Font.Size = 12
        Set-SubscriptCO2 $yAx.AxisTitle
        # Fixed 0–20 on every panel so the 2x2 sheet is comparable across measures
        # (auto-scaling gave 20/16/18/16 and made panels look misleadingly alike)
        $yAx.MinimumScale = 0; $yAx.MaximumScale = 20
        try { $yAx.TickLabels.Font.Name = "Sofia Pro" } catch { $yAx.TickLabels.Font.Name = "Calibri" }
        $yAx.TickLabels.Font.Size = 12
    } catch {}
    try {
        $ch.HasLegend = $true
        $ch.Legend.Position = -4107
        try { $ch.Legend.Font.Name = "Sofia Pro" } catch { $ch.Legend.Font.Name = "Calibri" }
        $ch.Legend.Font.Size = 12
    } catch {}
    # Plot layout LAST — title/series/legend edits above trigger a re-layout
    # that discards plot-area geometry set before them.
    # MUST be absolute, not relative: each chart sheet here is copied from the
    # previously formatted one, so a relative shift compounds with every copy
    # (+100, +200, +300, +400 — the 4th plot collapsed to 7 pt tall).
    # Canvas 870x520; legend sits at the BOTTOM (~489), so stop the plot short of it.
    try {
        $ch.PlotArea.InsideLeft   = 75
        $ch.PlotArea.InsideTop    = 55
        $ch.PlotArea.InsideWidth  = 770
        $ch.PlotArea.InsideHeight = 400
    } catch {}
}

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

Write-Host "`nFigure 9: Per-capita line chart (4 groups, 4 measures) ..."
$gCsv = Load-CSV "fig9_data.csv"
$wbG  = Open-Template "Chart-template-StackedArea.xlsx"

$wbG.Sheets("Data").Name  = "data_ghg_incLUC"
$wbG.Sheets("Chart").Name = "GHG incl LULUCF"

$wsG1    = $wbG.Sheets("data_ghg_incLUC")
$null    = Write-GData $wsG1 "ghg_incLUC" $gCsv "Annual GHG emissions including LULUCF per capita by country group, 1850-2024"
$chG1    = $wbG.Sheets("GHG incl LULUCF").ChartObjects(1).Chart
Set-GChart $chG1 $wsG1 "t CO2e per capita"

$wbG.Sheets("GHG incl LULUCF").Copy([System.Type]::Missing, $wbG.Sheets("GHG incl LULUCF"))
try { $wbG.ActiveSheet.Name = "CO2 excl LULUCF" } catch { $wbG.Sheets("GHG incl LULUCF (2)").Name = "CO2 excl LULUCF" }
$wsG2 = $wbG.Sheets.Add([System.Type]::Missing, $wbG.Sheets("GHG incl LULUCF"))
$wsG2.Name = "data_co2_excLUC"
$null = Write-GData $wsG2 "co2_excLUC" $gCsv "Annual CO2 emissions excluding LULUCF per capita by country group, 1850-2024"
$chG2 = $wbG.Sheets("CO2 excl LULUCF").ChartObjects(1).Chart
Set-GChart $chG2 $wsG2 "t CO2 per capita"

$wbG.Sheets("CO2 excl LULUCF").Copy([System.Type]::Missing, $wbG.Sheets("CO2 excl LULUCF"))
try { $wbG.ActiveSheet.Name = "GHG excl LULUCF" } catch { $wbG.Sheets("CO2 excl LULUCF (2)").Name = "GHG excl LULUCF" }
$wsG3 = $wbG.Sheets.Add([System.Type]::Missing, $wbG.Sheets("CO2 excl LULUCF"))
$wsG3.Name = "data_ghg_excLUC"
$null = Write-GData $wsG3 "ghg_excLUC" $gCsv "Annual GHG emissions excluding LULUCF per capita by country group, 1850-2024"
$chG3 = $wbG.Sheets("GHG excl LULUCF").ChartObjects(1).Chart
Set-GChart $chG3 $wsG3 "t CO2e per capita"

$wbG.Sheets("GHG excl LULUCF").Copy([System.Type]::Missing, $wbG.Sheets("GHG excl LULUCF"))
try { $wbG.ActiveSheet.Name = "CO2 incl LULUCF" } catch { $wbG.Sheets("GHG excl LULUCF (2)").Name = "CO2 incl LULUCF" }
$wsG4 = $wbG.Sheets.Add([System.Type]::Missing, $wbG.Sheets("GHG excl LULUCF"))
$wsG4.Name = "data_co2_incLUC"
$null = Write-GData $wsG4 "co2_incLUC" $gCsv "Annual CO2 emissions including LULUCF per capita by country group, 1850-2024"
$chG4 = $wbG.Sheets("CO2 incl LULUCF").ChartObjects(1).Chart
Set-GChart $chG4 $wsG4 "t CO2 per capita"

# Tab colours: data = dark teal, chart = amber (Figure_H.xlsx convention)
$_teal  = 5983243
$_amber = 2930175
foreach ($n in @("data_ghg_incLUC","data_co2_excLUC","data_ghg_excLUC","data_co2_incLUC")) {
    try { $wbG.Sheets($n).Tab.Color = $_teal  } catch {}
}
foreach ($n in @("GHG incl LULUCF","CO2 excl LULUCF","GHG excl LULUCF","CO2 incl LULUCF")) {
    try { $wbG.Sheets($n).Tab.Color = $_amber } catch {}
}

# Combined 2x2 sheet: GHGs on top row, incl. LULUCF in left column (no overall title)
$PANEL_TITLES  = @("All GHGs, including LULUCF", "All GHGs, excluding LULUCF", "CO2, including LULUCF", "CO2, excluding LULUCF")
Add-CombinedSheet $wbG @("GHG incl LULUCF","GHG excl LULUCF","CO2 incl LULUCF","CO2 excl LULUCF") "All measures" $PANEL_TITLES | Out-Null

Save-Close $wbG "$outDir\Figure_9.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
