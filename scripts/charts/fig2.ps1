<#
.SYNOPSIS
fig2.ps1
Generate Figure_2.xlsx -- CO2 + GHG combined stacked area (paper Figure 2).

Inputs:
  data/outputs/charts/fig2_stacked.csv     (from 09_prepare_chart_data.py)
  data/outputs/charts/fig2_combined.csv    (from 09_prepare_chart_data.py)
  cgd-general/chart-templates/Chart-template-StackedArea.xlsx

Output:
  data/outputs/charts/Figure_2.xlsx        (combined; tabs-only version not produced)
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

$CHART_TITLE_F2  = "Changing composition of CO2 and GHG emissions"
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
# Fig 2: StackedArea — OWID World, 1850-2024, two tabs (CO2 + GHGs)
# ---------------------------------------------------------------------------
Write-Host "`nFig 2: StackedArea (CO2 + GHGs, 2 tabs) ..."
$f2 = Load-CSV "fig2_stacked.csv"

$F2_MEASURES = @("CO2","GHG")
$F2_LABELS   = @{ "CO2"="CO2"; "GHG"="GHGs" }
$F2_SERIES   = @("Fossil fuels and industry", "LULUCF")

$wbF2 = Open-Template "Chart-template-StackedArea.xlsx"
$firstF2 = $F2_MEASURES[0]
$wbF2.Sheets("Data").Name  = "data_${firstF2}"
$wbF2.Sheets("Chart").Name = $F2_LABELS[$firstF2]

$extraF2 = @()
for ($mi = 1; $mi -lt $F2_MEASURES.Count; $mi++) {
    $msr2   = $F2_MEASURES[$mi]
    $tabF2  = $F2_LABELS[$msr2]
    $ndnF2  = "data_${msr2}"
    $pairF2 = Add-MeasureTab $wbF2 "data_${firstF2}" $F2_LABELS[$firstF2] $ndnF2 $tabF2
    $extraF2 += [PSCustomObject]@{ Ws=$pairF2[0]; Msr=$msr2; Label=$tabF2 }
}

function Fill-Fig2Tab($wsF2, $chartF2, $csvRows) {
    try { $wsF2.Cells.UnMerge() } catch {}
    $wsF2.Range("A3:H200").ClearContents()
    try { $wsF2.Range("D:H").Delete() | Out-Null } catch {}
    $wsF2.Cells(2,2).Value2 = $F2_SERIES[0]
    $wsF2.Cells(2,3).Value2 = $F2_SERIES[1]
    $r = 3
    foreach ($row in $csvRows) {
        $wsF2.Cells($r,1).Value2 = [int]$row.year
        Set-Value $wsF2.Cells($r,2) $row.fossil_industry_Mt
        Set-Value $wsF2.Cells($r,3) $row.lulucf_Mt
        $r++
    }
    $endR = $r - 1
    while ($chartF2.SeriesCollection().Count -gt 2) {
        $chartF2.SeriesCollection($chartF2.SeriesCollection().Count).Delete() | Out-Null
    }
    while ($chartF2.SeriesCollection().Count -lt 2) {
        $chartF2.SeriesCollection().NewSeries() | Out-Null
    }
    $dn = $wsF2.Name
    $chartF2.SeriesCollection(1).Name    = "='${dn}'!`$B`$2"
    $chartF2.SeriesCollection(1).Values  = $wsF2.Range("B3:B${endR}")
    $chartF2.SeriesCollection(1).XValues = $wsF2.Range("A3:A${endR}")
    $chartF2.SeriesCollection(2).Name    = "='${dn}'!`$C`$2"
    $chartF2.SeriesCollection(2).Values  = $wsF2.Range("C3:C${endR}")
    $chartF2.SeriesCollection(2).XValues = $wsF2.Range("A3:A${endR}")
    Set-ChartTitle $wsF2 $chartF2 $CHART_TITLE_F2
    Format-DataRange $wsF2 3 $endR 1 1 "0"
    Format-DataRange $wsF2 3 $endR 2 3 "#,##0"
    Format-Sheet $wsF2 3 $endR 3
    try {
        $xAx2 = $chartF2.Axes(1)
        $xAx2.TickLabelSpacing = 10
        $xAx2.TickMarkSpacing  = 10
        try { $xAx2.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx2.TickLabels.Font.Name = "Calibri" }
        $xAx2.TickLabels.Font.Size = 9
    } catch {}
    try {
        $yAx2 = $chartF2.Axes(2)
        $yAx2.TickLabels.NumberFormat = "0,"
        $yAx2.HasTitle = $true
        $yAx2.AxisTitle.Text = "Gt"
        try { $yAx2.AxisTitle.Font.Name = "Sofia Pro" } catch { $yAx2.AxisTitle.Font.Name = "Calibri" }
        $yAx2.AxisTitle.Font.Size = 9
    } catch {}
    return $endR
}

$rowsF2_1 = @($f2 | Where-Object { $_.measure -eq $firstF2 })
$wsF2_1   = $wbF2.Sheets("data_${firstF2}")
$chF2_1   = $wbF2.Sheets($F2_LABELS[$firstF2]).ChartObjects(1).Chart
Fill-Fig2Tab $wsF2_1 $chF2_1 $rowsF2_1 | Out-Null

foreach ($exF2 in $extraF2) {
    $rowsF2_m = @($f2 | Where-Object { $_.measure -eq $exF2.Msr })
    $chF2_m   = $wbF2.Sheets($exF2.Label).ChartObjects(1).Chart
    Fill-Fig2Tab $exF2.Ws $chF2_m $rowsF2_m | Out-Null
}

# Tabs-only version (Fig2_v4.xlsx) not produced — paper figure is the combined version below.
$wbF2.Close($false)

# ---------------------------------------------------------------------------
# Fig 2 combined: StackedArea — OWID World 1850-2024, 4 series
# ---------------------------------------------------------------------------
Write-Host "`nFig 2 combined: StackedArea (4 series) ..."
$f2c = Load-CSV "fig2_combined.csv"

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
