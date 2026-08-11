<#
.SYNOPSIS
figE.ps1
Generate Figure_7.xlsx -- reverse cumulative share MultiLine (4 measures).

Inputs:
  data/outputs/charts/fig7_multiline.csv   (from 09_prepare_chart_data.py)
  cgd-general/chart-templates/Chart-template-MultiLine.xlsx

Output:
  data/outputs/charts/Figure_7.xlsx

Pre-req: scripts 01, 07, 09 must have run.
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

$ML_COLS = @{
    "fig7" = @{
        "co2_excLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_GCP_fossil","");                B = @("annex1_OWID","annex1_PRIMAP","annex1_GCP_fossil","");                C = @("non_annex1_OWID","non_annex1_PRIMAP","non_annex1_GCP_fossil","") }
        "co2_incLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_GCP_BLUE","annex2_GCP_OSCAR"); B = @("annex1_OWID","annex1_PRIMAP","annex1_GCP_BLUE","annex1_GCP_OSCAR"); C = @("non_annex1_OWID","non_annex1_PRIMAP","non_annex1_GCP_BLUE","non_annex1_GCP_OSCAR") }
        "ghg_excLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_EDGAR","annex2_ClimateWatch"); B = @("annex1_OWID","annex1_PRIMAP","annex1_EDGAR","annex1_ClimateWatch"); C = @("non_annex1_OWID","non_annex1_PRIMAP","non_annex1_EDGAR","non_annex1_ClimateWatch") }
        "ghg_incLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_EDGAR","");                    B = @("annex1_OWID","annex1_PRIMAP","annex1_EDGAR","");                    C = @("non_annex1_OWID","non_annex1_PRIMAP","non_annex1_EDGAR","") }
    }
}

$ML_TITLES = @{
    "fig7" = @{
        "co2_excLUC" = "Cumulative CO2 emissions excluding LULUCF to 2024, different start points"
        "co2_incLUC" = "Cumulative CO2 emissions including from LULUCF to 2024, different start points"
        "ghg_excLUC" = "Cumulative GHG emissions excluding LULUCF to 2024, different start points"
        "ghg_incLUC" = "Cumulative GHG emissions including LULUCF to 2024, different start points"
    }
}

. "$PSScriptRoot\_chart_helpers.ps1"

$cgdOlive = 4489605

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

. "$PSScriptRoot\_chart_multiline.ps1"

Write-Host "`nFig 7: MultiLine (4 measures) ..."

$figKey = "fig7"
$mlCsv  = Load-CSV "${figKey}_multiline.csv"
$wbML   = Open-Template "Chart-template-MultiLine.xlsx"

$fMsr = $MEASURES[0]
$wbML.Sheets("Data").Name = "data_${fMsr}"
try { $wbML.Sheets("Chart_colours").Delete() } catch {}
$wbML.Sheets("Chart_dashes").Name = $MLABEL[$fMsr]

$ws1Vals = $wbML.Sheets.Add()
$ws1Vals.Name = "vals_${fMsr}"

$extraML = @()
for ($mi = 1; $mi -lt $MEASURES.Count; $mi++) {
    $msr    = $MEASURES[$mi]
    $ndn    = "data_${msr}"
    $pairML = Add-MeasureTab $wbML "data_${fMsr}" $MLABEL[$fMsr] $ndn $MLABEL[$msr]
    $wsValsM = $wbML.Sheets.Add()
    $wsValsM.Name = "vals_${msr}"
    $extraML += [PSCustomObject]@{ WsData=$pairML[0]; WsVals=$wsValsM; Msr=$msr }
}

$cm1     = $ML_COLS[$figKey][$fMsr]
$ws1Data = $wbML.Sheets("data_${fMsr}")
$rows1ML = @($mlCsv | Where-Object { $_.measure -eq $fMsr })
$end1Vals = Build-MultiLineVals $ws1Vals "vals_${fMsr}" $figKey $rows1ML $cm1
Build-MultiLineCalcs $ws1Data "vals_${fMsr}" $cm1 $end1Vals $rows1ML

foreach ($exML in $extraML) {
    $cmM    = $ML_COLS[$figKey][$exML.Msr]
    $rowsML = @($mlCsv | Where-Object { $_.measure -eq $exML.Msr })
    $endV   = Build-MultiLineVals $exML.WsVals "vals_$($exML.Msr)" $figKey $rowsML $cmM
    Build-MultiLineCalcs $exML.WsData "vals_$($exML.Msr)" $cmM $endV $rowsML
}

$allMLDataSheets = @("data_${fMsr}") + ($extraML | ForEach-Object { "data_" + $_.Msr })
$allMLChartTabs  = @($MLABEL[$fMsr]) + ($extraML | ForEach-Object { $MLABEL[$_.Msr] })
$allMLMsrs       = @($fMsr)          + ($extraML | ForEach-Object { $_.Msr })

for ($ti = 0; $ti -lt $allMLDataSheets.Count; $ti++) {
    $wsT = $wbML.Sheets($allMLDataSheets[$ti])
    $chT = $wbML.Sheets($allMLChartTabs[$ti]).ChartObjects(1).Chart
    $mT  = $allMLMsrs[$ti]

    Set-ChartTitle $wsT $chT $ML_TITLES[$figKey][$mT]

    try {
        $xAxML = $chT.Axes(1); $xAxML.TickLabelSpacing = 10; $xAxML.TickMarkSpacing = 10
        try { $xAxML.TickLabels.Font.Name = "Sofia Pro" } catch { $xAxML.TickLabels.Font.Name = "Calibri" }
        $xAxML.TickLabels.Font.Size = 12
    } catch {}
    try {
        $yAx = $chT.Axes(2)
        $yAx.HasTitle = $true
        $yAx.AxisTitle.Text = "Cumulative share of emissions (%)"
        $yAx.MinimumScale = 0
        $yAx.MaximumScale = 100
        try { $yAx.AxisTitle.Font.Name = "Sofia Pro" } catch { $yAx.AxisTitle.Font.Name = "Calibri" }
        $yAx.AxisTitle.Font.Size = 12
        try { $yAx.TickLabels.Font.Name = "Sofia Pro" } catch { $yAx.TickLabels.Font.Name = "Calibri" }
        $yAx.TickLabels.Font.Size = 12
    } catch {}
    try { $chT.SeriesCollection(3).Format.Fill.ForeColor.RGB = $cgdOlive } catch {}
    try { $chT.SeriesCollection(3).Format.Fill.Transparency  = 0.5 } catch {}
    try { $chT.SeriesCollection(3).Format.Line.Visible       = $false } catch {}
    try { $chT.SeriesCollection(2).Format.Fill.Transparency  = 0.7 } catch {}
    try { $chT.SeriesCollection(4).Format.Fill.Transparency  = 0.99 } catch {}
    try { $chT.SeriesCollection(5).Format.Fill.Transparency  = 0.5  } catch {}

    $lastR = $wsT.Cells($wsT.Rows.Count, 1).End(-4162).Row
    if ($lastR -ge 3) {
        Add-GMSTLines $chT $wsT 3 $lastR
        Format-DataRange $wsT 3 $lastR 1 1 "0"
        Format-DataRange $wsT 3 $lastR 9 19 "0.0"
        Format-Sheet $wsT 3 $lastR 19
    }

    try {
        try { $chT.Legend.Font.Name = "Sofia Pro" } catch { $chT.Legend.Font.Name = "Calibri" }
        $chT.Legend.Font.Size = 12
        $nLe = $chT.Legend.LegendEntries().Count
        for ($le = $nLe; $le -ge 1; $le--) {
            if ($le -ne 1 -and $le -ne 3) { try { $chT.Legend.LegendEntries($le).Delete() } catch {} }
        }
    } catch {}

    # Plot layout LAST — series/legend edits above trigger re-layout that
    # discards earlier plot-area settings. Fill the 870x520 canvas: small
    # right/bottom margins, room for title+legend above and x labels below.
    try {
        $chT.PlotArea.InsideLeft   = 75
        $chT.PlotArea.InsideTop    = 100
        $chT.PlotArea.InsideWidth  = 770
        $chT.PlotArea.InsideHeight = 375
    } catch {}
}

# Combined 2x2 sheet: GHGs on top row, incl. LULUCF in left column (no overall title)
$COMBINED_TABS  = @($MLABEL["ghg_incLUC"], $MLABEL["ghg_excLUC"], $MLABEL["co2_incLUC"], $MLABEL["co2_excLUC"])
$PANEL_TITLES   = @("All GHGs, including LULUCF", "All GHGs, excluding LULUCF", "CO2, including LULUCF", "CO2, excluding LULUCF")
Add-CombinedSheet $wbML $COMBINED_TABS "All measures" $PANEL_TITLES | Out-Null

Save-Close $wbML "$outDir\Figure_7.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
