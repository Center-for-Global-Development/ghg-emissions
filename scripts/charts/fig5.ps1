<#
.SYNOPSIS
figC.ps1
Generate Figure_5.xlsx -- total GHG emissions by country group, stacked area.

Inputs:
  data/outputs/charts/fig5_stacked.csv    (from 09_prepare_chart_data.py)
  cgd-general/chart-templates/Chart-template-StackedArea.xlsx

Output:
  data/outputs/charts/Figure_5.xlsx
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

$MEASURES = @("co2_excLUC","co2_incLUC","ghg_excLUC","ghg_incLUC")
$MLABEL   = @{
    "co2_excLUC" = "CO2 excl LULUCF"
    "co2_incLUC" = "CO2 incl LULUCF"
    "ghg_excLUC" = "GHG excl LULUCF"
    "ghg_incLUC" = "GHG incl LULUCF"
}

$CHART_TITLE_C = "Total GHG emissions, 1850 - 2024"

$wcGroups = @("usa","euro_a2","other_a2","other_ai","chn","ind","nai_g20","other_non_ai")
$wcLabels = @{
    "usa"          = "USA"
    "euro_a2"      = "European Annex 2"
    "other_a2"     = "Other Annex 2"
    "other_ai"     = "Other Annex I"
    "chn"          = "China"
    "ind"          = "India"
    "nai_g20"      = "Other non-Annex I G20 members"
    "other_non_ai" = "Other non-Annex I"
}

$FIGC_COLORS = @{
    "usa"          = 5983243
    "euro_a2"      = 10390042
    "other_a2"     = 13679965
    "other_ai"     = 14405518
    "chn"          = 2930175
    "ind"          = 36556
    "nai_g20"      = 5688063
    "other_non_ai" = 8443391
}

. "$PSScriptRoot\_chart_helpers.ps1"

function Fill-StackedArea($ws3, $csvRows, $groups, $startRow3) {
    $r = $startRow3
    foreach ($row in $csvRows) {
        $ws3.Cells($r,1).Value2 = [int]$row.year
        for ($c = 0; $c -lt $groups.Count; $c++) {
            Set-Value $ws3.Cells($r, $c + 2) $row.($groups[$c])
        }
        $r++
    }
    return $r - 1
}

function Set-StackedSeries($chart3, $ws3, $startRow3, $endRow3, $groups, $labels) {
    while ($chart3.SeriesCollection().Count -gt $groups.Count) {
        $chart3.SeriesCollection($chart3.SeriesCollection().Count).Delete() | Out-Null
    }
    while ($chart3.SeriesCollection().Count -lt $groups.Count) {
        $chart3.SeriesCollection().NewSeries() | Out-Null
    }
    $dn3 = $ws3.Name
    for ($i = 1; $i -le $groups.Count; $i++) {
        $s3 = $chart3.SeriesCollection($i)
        $c  = [char](65 + $i)
        $ws3.Cells(2, $i + 1).Value2 = $labels[$groups[$i-1]]
        $s3.Name    = "='${dn3}'!`$${c}`$2"
        $s3.Values  = $ws3.Range("${c}${startRow3}:${c}${endRow3}")
        $s3.XValues = $ws3.Range("A${startRow3}:A${endRow3}")
    }
}

function Set-FigCFormatting($chart3, $groups) {
    for ($i = 1; $i -le $groups.Count; $i++) {
        $grpKey = $groups[$i - 1]
        $color  = $FIGC_COLORS[$grpKey]
        if ($null -ne $color) {
            try { $chart3.SeriesCollection($i).Format.Fill.ForeColor.RGB = $color } catch {}
            # No series borders — bands read as flat colour blocks
            try { $chart3.SeriesCollection($i).Format.Line.Visible = $false } catch {}
            try { $chart3.SeriesCollection($i).Border.LineStyle = -4142 } catch {}   # xlNone
        }
    }
    try {
        $xAx = $chart3.Axes(1)
        $xAx.TickLabelSpacing = 10
        $xAx.TickMarkSpacing  = 10
        try { $xAx.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx.TickLabels.Font.Name = "Calibri" }
        $xAx.TickLabels.Font.Size = 12
    } catch {}
    try {
        $yAx = $chart3.Axes(2)
        $yAx.TickLabels.NumberFormat = "0,"
        # Fixed 0-60 Gt on every panel so the 2x2 sheet is comparable across
        # measures (auto gave 60/50/45/40; tallest stack is ghg_incLUC at 54.4 Gt).
        # Values are Mt; the "0," format displays them as Gt.
        $yAx.MinimumScale = 0
        $yAx.MaximumScale = 60000
        $yAx.MajorUnit    = 10000
        $yAx.HasTitle = $true
        $yAx.AxisTitle.Text = "Gt CO2e"
        try { $yAx.AxisTitle.Font.Name = "Sofia Pro" } catch { $yAx.AxisTitle.Font.Name = "Calibri" }
        $yAx.AxisTitle.Font.Size = 12
        Set-SubscriptCO2 $yAx.AxisTitle
        try { $yAx.TickLabels.Font.Name = "Sofia Pro" } catch { $yAx.TickLabels.Font.Name = "Calibri" }
        $yAx.TickLabels.Font.Size = 12
    } catch {}
    try {
        $chart3.HasLegend = $true
        try { $chart3.Legend.Font.Name = "Sofia Pro" } catch { $chart3.Legend.Font.Name = "Calibri" }
        $chart3.Legend.Font.Size = 12
    } catch {}
}

# Plot layout — must run LAST for a chart. Title, series and legend edits all
# trigger an Excel re-layout that discards plot-area geometry set before them.
# Absolute values (not relative +/-) so repeat calls are idempotent.
# Canvas is 870x520; legend sits at the top, so leave ~100 pt above the plot.
function Set-FigCPlotArea($chart3) {
    try {
        $chart3.PlotArea.InsideLeft   = 75
        $chart3.PlotArea.InsideTop    = 100
        $chart3.PlotArea.InsideWidth  = 770
        $chart3.PlotArea.InsideHeight = 375
    } catch {}
}

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

Write-Host "`nFig 5: StackedArea (8 country groups, 4 measures) ..."
$fC = Load-CSV "fig5_stacked.csv"

$wbC = Open-Template "Chart-template-StackedArea.xlsx"
$firstMsr = $MEASURES[0]
$wbC.Sheets("Data").Name  = "data_${firstMsr}"
$wbC.Sheets("Chart").Name = $MLABEL[$firstMsr]

$extraC = @()
for ($mi = 1; $mi -lt $MEASURES.Count; $mi++) {
    $msr     = $MEASURES[$mi]
    $tabName = $MLABEL[$msr]
    $ndn     = "data_${msr}"
    $pair    = Add-MeasureTab $wbC "data_${firstMsr}" $MLABEL[$firstMsr] $ndn $tabName
    $extraC += [PSCustomObject]@{ Ws=$pair[0]; Msr=$msr; Label=$tabName }
}

$rows1C = @($fC | Where-Object { $_.measure -eq $firstMsr })
$ws1C   = $wbC.Sheets("data_${firstMsr}")
try { $ws1C.Cells.UnMerge() } catch {}
$ws1C.Range("A3:I200").ClearContents()
$end1C  = Fill-StackedArea $ws1C $rows1C $wcGroups 3
$ch1C   = $wbC.Sheets($MLABEL[$firstMsr]).ChartObjects(1).Chart
Set-StackedSeries $ch1C $ws1C 3 $end1C $wcGroups $wcLabels
Set-FigCFormatting $ch1C $wcGroups
Set-ChartTitle $ws1C $ch1C $CHART_TITLE_C
Format-DataRange $ws1C 3 $end1C 1 1 "0"
Format-DataRange $ws1C 3 $end1C 2 9 "#,##0"
Format-Sheet $ws1C 3 $end1C 9
Set-FigCPlotArea $ch1C

foreach ($exC in $extraC) {
    try { $exC.Ws.Cells.UnMerge() } catch {}
    $exC.Ws.Range("A3:I200").ClearContents()
    $rowsM  = @($fC | Where-Object { $_.measure -eq $exC.Msr })
    $endMC  = Fill-StackedArea $exC.Ws $rowsM $wcGroups 3
    $newChC = $wbC.Sheets($exC.Label).ChartObjects(1).Chart
    Set-StackedSeries $newChC $exC.Ws 3 $endMC $wcGroups $wcLabels
    Set-FigCFormatting $newChC $wcGroups
    Set-ChartTitle $exC.Ws $newChC $CHART_TITLE_C
    Format-DataRange $exC.Ws 3 $endMC 1 1 "0"
    Format-DataRange $exC.Ws 3 $endMC 2 9 "#,##0"
    Format-Sheet $exC.Ws 3 $endMC 9
    Set-FigCPlotArea $newChC
}

# Combined 2x2 sheet: GHGs on top row, incl. LULUCF in left column (no overall title)
$COMBINED_TABS   = @($MLABEL["ghg_incLUC"], $MLABEL["ghg_excLUC"], $MLABEL["co2_incLUC"], $MLABEL["co2_excLUC"])
$PANEL_TITLES    = @("All GHGs, including LULUCF", "All GHGs, excluding LULUCF", "CO2, including LULUCF", "CO2, excluding LULUCF")
Add-CombinedSheet $wbC $COMBINED_TABS "All measures" $PANEL_TITLES | Out-Null

Save-Close $wbC "$outDir\Figure_5.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
