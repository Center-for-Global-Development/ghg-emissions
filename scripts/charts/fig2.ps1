<#
.SYNOPSIS
fig2.ps1
Generate Figure_2.xlsx -- CO2 + GHG combined stacked area (paper Figure 2).

Inputs:
  data/outputs/charts/fig2_data.csv    (from 09_prepare_chart_data.py)
  templates/Chart-template-StackedArea.xlsx

Output:
  data/outputs/charts/Figure_2.xlsx
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

$cgdTeal      = 5983243
$cgdAmber     = 2930175
$cgdTealMid   = 10390042
$cgdAmberDark = 36556

$CHART_TITLE_F2C = "Composition of total GHG emissions by gas type and source, 1850-2024"

. "$PSScriptRoot\_chart_helpers.ps1"

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {
# ---------------------------------------------------------------------------
# Fig 2 combined: StackedArea — OWID World 1850-2024, 4 series
# ---------------------------------------------------------------------------
Write-Host "`nFig 2 combined: StackedArea (4 series) ..."
$f2c = Load-CSV "fig2_data.csv"

$F2C_SERIES = @(
    "CO2 - fossil fuels and industry",
    "Non-CO2 GHGs - fossil fuels and industry",
    "CO2 - LULUCF",
    "Non-CO2 GHGs - LULUCF"
)
$F2C_COLS   = @("co2_fossil_Mt","nonco2_fossil_Mt","co2_lulucf_Mt","nonco2_lulucf_Mt")
$F2C_COLORS = @($cgdTeal, $cgdTealMid, $cgdAmberDark, $cgdAmber)

$wbF2c = Open-Template "Chart-template-StackedArea.xlsx"
$wsF2c = $wbF2c.Sheets("Data")
$chF2c = $wbF2c.Sheets("Chart").ChartObjects(1).Chart

try { $wsF2c.Cells.UnMerge() } catch {}
$wsF2c.Range("A1:H200").ClearContents()
try { $wsF2c.Range("F:H").Delete() | Out-Null } catch {}

for ($si = 0; $si -lt $F2C_SERIES.Count; $si++) {
    $wsF2c.Cells(2, $si + 2).Value2 = $F2C_SERIES[$si]
}

$r = 3
foreach ($row in $f2c) {
    $wsF2c.Cells($r, 1).Value2 = [int]$row.year
    for ($si = 0; $si -lt $F2C_COLS.Count; $si++) {
        Set-Value $wsF2c.Cells($r, $si + 2) $row.($F2C_COLS[$si])
    }
    $r++
}
$endR2c = $r - 1

while ($chF2c.SeriesCollection().Count -gt 4) {
    $chF2c.SeriesCollection($chF2c.SeriesCollection().Count).Delete() | Out-Null
}
while ($chF2c.SeriesCollection().Count -lt 4) {
    $chF2c.SeriesCollection().NewSeries() | Out-Null
}

$dn2c = $wsF2c.Name
for ($si = 1; $si -le 4; $si++) {
    $colLetter = [char](64 + $si + 1)
    $chF2c.SeriesCollection($si).Name    = "='${dn2c}'!`$${colLetter}`$2"
    $chF2c.SeriesCollection($si).Values  = $wsF2c.Range("${colLetter}3:${colLetter}${endR2c}")
    $chF2c.SeriesCollection($si).XValues = $wsF2c.Range("A3:A${endR2c}")
    try { $chF2c.SeriesCollection($si).Interior.Color = $F2C_COLORS[$si - 1] } catch {}
}

# Title kept in Data!A1 only; no title on the chart itself
$wsF2c.Cells(1,1).Value2 = $CHART_TITLE_F2C
try { $wsF2c.Range("A1:E1").Merge() | Out-Null } catch {}
$chF2c.HasTitle = $false

try {
    $xAx2c = $chF2c.Axes(1)
    $xAx2c.TickLabelSpacing = 10
    $xAx2c.TickMarkSpacing  = 10
    try { $xAx2c.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx2c.TickLabels.Font.Name = "Calibri" }
    $xAx2c.TickLabels.Font.Size = 12
} catch {}
try {
    $yAx2c = $chF2c.Axes(2)
    $yAx2c.TickLabels.NumberFormat = "0,"
    $yAx2c.HasTitle = $true
    $yAx2c.AxisTitle.Text = "Gt"
    try { $yAx2c.AxisTitle.Font.Name = "Sofia Pro" } catch { $yAx2c.AxisTitle.Font.Name = "Calibri" }
    $yAx2c.AxisTitle.Font.Size = 12
    try { $yAx2c.TickLabels.Font.Name = "Sofia Pro" } catch { $yAx2c.TickLabels.Font.Name = "Calibri" }
    $yAx2c.TickLabels.Font.Size = 12
} catch {}
try {
    try { $chF2c.Legend.Font.Name = "Sofia Pro" } catch { $chF2c.Legend.Font.Name = "Calibri" }
    $chF2c.Legend.Font.Size = 12
} catch {}

Format-DataRange $wsF2c 3 $endR2c 1 1 "0"
Format-DataRange $wsF2c 3 $endR2c 2 5 "#,##0"
Format-Sheet $wsF2c 3 $endR2c 5

Save-Close $wbF2c "$outDir\Figure_2.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
