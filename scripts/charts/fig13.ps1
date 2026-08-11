<#
.SYNOPSIS
figK.ps1
Generate Figure_13.xlsx -- annual % share of world emissions, MultiLine (4 measures).

Inputs:
  data/outputs/charts/fig13_multiline.csv   (from 09_prepare_chart_data.py)
  cgd-general/chart-templates/Chart-template-MultiLine.xlsx

Output:
  data/outputs/charts/Figure_13.xlsx

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
    "fig13" = @{
        "co2_excLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_GCP_fossil","");               B = @("annex1_OWID","annex1_PRIMAP","annex1_GCP_fossil","") }
        "co2_incLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_GCP_BLUE","annex2_GCP_OSCAR"); B = @("annex1_OWID","annex1_PRIMAP","annex1_GCP_BLUE","annex1_GCP_OSCAR") }
        "ghg_excLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_EDGAR","annex2_ClimateWatch"); B = @("annex1_OWID","annex1_PRIMAP","annex1_EDGAR","annex1_ClimateWatch") }
        "ghg_incLUC" = @{ A = @("annex2_OWID","annex2_PRIMAP","annex2_EDGAR","");                    B = @("annex1_OWID","annex1_PRIMAP","annex1_EDGAR","") }
    }
}

$ML_TITLES = @{
    "fig13" = @{
        "co2_excLUC" = "Annual CO2 emissions excluding LULUCF, share of world total, 1850-2024"
        "co2_incLUC" = "Annual CO2 emissions including LULUCF, share of world total, 1850-2024"
        "ghg_excLUC" = "Annual GHG emissions excluding LULUCF, share of world total, 1850-2024"
        "ghg_incLUC" = "Annual GHG emissions including LULUCF, share of world total, 1850-2024"
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

Write-Host "`nFig 13: MultiLine annual % shares (4 measures) ..."

$figKey = "fig13"
$klCsv  = Load-CSV "fig13_multiline.csv"
$wbKL   = Open-Template "Chart-template-MultiLine.xlsx"

$fMsr = $MEASURES[0]
$wbKL.Sheets("Data").Name = "data_${fMsr}"
try { $wbKL.Sheets("Chart_colours").Delete() } catch {}
$wbKL.Sheets("Chart_dashes").Name = $MLABEL[$fMsr]

$extraKL = @()
for ($mi = 1; $mi -lt $MEASURES.Count; $mi++) {
    $msr    = $MEASURES[$mi]
    $ndn    = "data_${msr}"
    $pairKL = Add-MeasureTab $wbKL "data_${fMsr}" $MLABEL[$fMsr] $ndn $MLABEL[$msr]
    $extraKL += [PSCustomObject]@{ WsData=$pairKL[0]; Msr=$msr }
}

$cmK1   = $ML_COLS[$figKey][$fMsr]
$ws1K   = $wbKL.Sheets("data_${fMsr}")
$rows1K = @($klCsv | Where-Object { $_.measure -eq $fMsr })
$null   = Build-FigKData $ws1K $cmK1 $rows1K

foreach ($exKL in $extraKL) {
    $cmM   = $ML_COLS[$figKey][$exKL.Msr]
    $rowsM = @($klCsv | Where-Object { $_.measure -eq $exKL.Msr })
    $null  = Build-FigKData $exKL.WsData $cmM $rowsM
}

$allKData  = @("data_${fMsr}") + ($extraKL | ForEach-Object { "data_" + $_.Msr })
$allKChart = @($MLABEL[$fMsr]) + ($extraKL | ForEach-Object { $MLABEL[$_.Msr] })
$allKMsrs  = @($fMsr)          + ($extraKL | ForEach-Object { $_.Msr })

for ($ti = 0; $ti -lt $allKData.Count; $ti++) {
    $wsT = $wbKL.Sheets($allKData[$ti])
    $chT = $wbKL.Sheets($allKChart[$ti]).ChartObjects(1).Chart
    $mT  = $allKMsrs[$ti]

    Set-ChartTitle $wsT $chT $ML_TITLES[$figKey][$mT]
    try {
        $xAxK = $chT.Axes(1); $xAxK.TickLabelSpacing = 10; $xAxK.TickMarkSpacing = 10
        try { $xAxK.TickLabels.Font.Name = "Sofia Pro" } catch { $xAxK.TickLabels.Font.Name = "Calibri" }
        $xAxK.TickLabels.Font.Size = 12
    } catch {}
    try {
        $yAxK = $chT.Axes(2)
        $yAxK.HasTitle = $true
        $yAxK.AxisTitle.Text = "Annual emissions share (%)"
        $yAxK.MinimumScale = 0
        $yAxK.MaximumScale = 100
        try { $yAxK.AxisTitle.Font.Name = "Sofia Pro" } catch { $yAxK.AxisTitle.Font.Name = "Calibri" }
        $yAxK.AxisTitle.Font.Size = 12
        try { $yAxK.TickLabels.Font.Name = "Sofia Pro" } catch { $yAxK.TickLabels.Font.Name = "Calibri" }
        $yAxK.TickLabels.Font.Size = 12
    } catch {}
    try { $chT.SeriesCollection(3).Format.Fill.ForeColor.RGB = $cgdOlive } catch {}
    try { $chT.SeriesCollection(3).Format.Fill.Transparency  = 0.5 } catch {}
    try { $chT.SeriesCollection(3).Format.Line.Visible       = $false } catch {}
    try { $chT.SeriesCollection(2).Format.Fill.Transparency  = 0.7 } catch {}
    try { $chT.SeriesCollection(4).Format.Fill.Transparency  = 0.99 } catch {}
    try { $chT.SeriesCollection(5).Format.Fill.Transparency  = 0.5  } catch {}

    $lastRK = $wsT.Cells($wsT.Rows.Count, 1).End(-4162).Row
    if ($lastRK -ge 3) {
        Format-DataRange $wsT 3 $lastRK 1 1 "0"
        Format-DataRange $wsT 3 $lastRK 9 17 "0.0"
        Format-Sheet $wsT 3 $lastRK 17
    }
    try {
        try { $chT.Legend.Font.Name = "Sofia Pro" } catch { $chT.Legend.Font.Name = "Calibri" }
        $chT.Legend.Font.Size = 12
        $nLeK = $chT.Legend.LegendEntries().Count
        for ($le = $nLeK; $le -ge 1; $le--) {
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
Add-CombinedSheet $wbKL $COMBINED_TABS "All measures" $PANEL_TITLES | Out-Null

Save-Close $wbKL "$outDir\Figure_13.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
