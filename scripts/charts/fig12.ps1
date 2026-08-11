<#
.SYNOPSIS
figJ.ps1
Generate Figure_12.xlsx -- annual % share of GHG incl. LUC (single-tab MultiLine).

Inputs:
  data/outputs/charts/fig12_multiline.csv  (ghg_incLUC rows; built by script 09)
  cgd-general/chart-templates/Chart-template-MultiLine.xlsx

Output:
  data/outputs/charts/Figure_12.xlsx

Pre-req: script 09 must have run to produce fig12_multiline.csv.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

. "$PSScriptRoot\_chart_helpers.ps1"

$cgdTealMid = 10390042
$cgdOlive   = 4489605

# Source slot mapping for ghg_incLUC (OWID + PRIMAP; EDGAR blank; CW blank)
$cmJ = @{
    A = @("annex2_OWID","annex2_PRIMAP","","")
    B = @("annex1_OWID","annex1_PRIMAP","","")
}

function Set-Value-Local { param($cell, $val)
    if ([string]::IsNullOrWhiteSpace($val) -or $val -eq "NaN") { return }
    try { $cell.Value2 = [double]$val } catch {}
}

function Build-FigKData($wsD, $cm, $csvRows) {
    try { $null = $wsD.Cells.UnMerge() } catch {}
    $null = $wsD.Range("A2:S200").ClearContents()
    $endRowInt = 2 + @($csvRows).Count
    if ($endRowInt -lt 3) { return $endRowInt }
    $wsD.Cells(2,  1).Value2 = "Year"
    $wsD.Cells(2,  2).Value2 = "AII_Min"
    $wsD.Cells(2,  3).Value2 = "AII_Max"
    $wsD.Cells(2,  4).Value2 = "AI_Min"
    $wsD.Cells(2,  5).Value2 = "AI_Max"
    $wsD.Cells(2,  6).Value2 = "Pure_AII"
    $wsD.Cells(2,  7).Value2 = "Overlap"
    $wsD.Cells(2,  8).Value2 = "AI_Gap"
    $wsD.Cells(2, 19).Value2 = "Pure_AI"
    for ($i = 0; $i -lt 4; $i++) {
        $a2col = $cm.A[$i]
        if ($a2col -eq "") { continue }
        $src = $a2col -replace "^annex2_", ""
        $wsD.Cells(2,  9 + $i).Value2 = "pct_annex2_${src}"
        $wsD.Cells(2, 13 + $i).Value2 = "pct_annex1_${src}"
    }
    $r = 3
    foreach ($row in $csvRows) {
        $wsD.Cells($r, 1).Value2 = [int]$row.year
        for ($i = 0; $i -lt 4; $i++) {
            $a2col = $cm.A[$i]; $a1col = $cm.B[$i]
            if ($a2col -ne "") { Set-Value-Local $wsD.Cells($r, 9+$i) $row.$a2col }
            if ($a1col -ne "") { Set-Value-Local $wsD.Cells($r, 13+$i) $row.$a1col }
        }
        $r++
    }
    $wsD.Range($wsD.Cells(3,2), $wsD.Cells($endRowInt,2)).FormulaR1C1 = "=IFERROR(MIN(RC[7]:RC[10]),"""")"
    $wsD.Range($wsD.Cells(3,3), $wsD.Cells($endRowInt,3)).FormulaR1C1 = "=IFERROR(MAX(RC[6]:RC[9]),"""")"
    $wsD.Range($wsD.Cells(3,4), $wsD.Cells($endRowInt,4)).FormulaR1C1 = "=IFERROR(MIN(RC[9]:RC[12]),"""")"
    $wsD.Range($wsD.Cells(3,5), $wsD.Cells($endRowInt,5)).FormulaR1C1 = "=IFERROR(MAX(RC[8]:RC[11]),"""")"
    $wsD.Range($wsD.Cells(3,6), $wsD.Cells($endRowInt,6)).FormulaR1C1 = "=IFERROR(MIN(RC[-3],RC[-2])-RC[-4],"""")"
    $wsD.Range($wsD.Cells(3,7), $wsD.Cells($endRowInt,7)).FormulaR1C1 = "=IFERROR(MAX(0,RC[-4]-RC[-3]),"""")"
    $wsD.Range($wsD.Cells(3,8), $wsD.Cells($endRowInt,8)).FormulaR1C1 = "=IFERROR(MAX(0,RC[-4]-RC[-5]),"""")"
    $wsD.Range($wsD.Cells(3,19), $wsD.Cells($endRowInt,19)).FormulaR1C1 = "=IFERROR(MAX(0,RC5-MAX(RC3,RC4)),"""")"
    return $endRowInt
}

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

Write-Host "`nFigure 12: Annual GHG incl. LUC % shares ..."
$fkCsv = Load-CSV "fig12_multiline.csv"
$wbJ   = Open-Template "Chart-template-MultiLine.xlsx"

$wbJ.Sheets("Data").Name = "data_ghg_incLUC"
try { $wbJ.Sheets("Chart_colours").Delete() } catch {}
$wbJ.Sheets("Chart_dashes").Name = "GHG incl LUC"

$wsJ   = $wbJ.Sheets("data_ghg_incLUC")
$rowsJ = @($fkCsv | Where-Object { $_.measure -eq "ghg_incLUC" })
$null  = Build-FigKData $wsJ $cmJ $rowsJ
$chJ   = $wbJ.Sheets("GHG incl LUC").ChartObjects(1).Chart

# Title kept in data sheet A1 only; no title on the chart itself
$wsJ.Cells(1,1).Value2 = "Annual share of all GHG emissions including LULUCF (%)"
$chJ.HasTitle = $false

try {
    $xAxJ = $chJ.Axes(1); $xAxJ.TickLabelSpacing = 10; $xAxJ.TickMarkSpacing = 10
    try { $xAxJ.TickLabels.Font.Name = "Sofia Pro" } catch { $xAxJ.TickLabels.Font.Name = "Calibri" }
    $xAxJ.TickLabels.Font.Size = 12
} catch {}
try {
    $yAxJ = $chJ.Axes(2)
    $yAxJ.HasTitle = $true
    $yAxJ.AxisTitle.Text = "Annual emissions share (%)"
    $yAxJ.MinimumScale = 0
    $yAxJ.MaximumScale = 100
    try { $yAxJ.AxisTitle.Font.Name = "Sofia Pro" } catch { $yAxJ.AxisTitle.Font.Name = "Calibri" }
    $yAxJ.AxisTitle.Font.Size = 12
    try { $yAxJ.TickLabels.Font.Name = "Sofia Pro" } catch { $yAxJ.TickLabels.Font.Name = "Calibri" }
    $yAxJ.TickLabels.Font.Size = 12
} catch {}
try {
    $chJ.PlotArea.InsideLeft   = 100; $chJ.PlotArea.InsideTop    = 80
    $chJ.PlotArea.InsideWidth  = 560; $chJ.PlotArea.InsideHeight = 380
} catch {}

try { $chJ.SeriesCollection(3).Format.Fill.ForeColor.RGB = $cgdOlive } catch {}
try { $chJ.SeriesCollection(3).Format.Fill.Transparency  = 0.5 } catch {}
try { $chJ.SeriesCollection(3).Format.Line.Visible       = $false } catch {}
try { $chJ.SeriesCollection(2).Format.Fill.Transparency  = 0.7 } catch {}
try { $chJ.SeriesCollection(4).Format.Fill.Transparency  = 0.99 } catch {}

$lastRJ = $wsJ.Cells($wsJ.Rows.Count, 1).End(-4162).Row
if ($lastRJ -ge 3) {
    $sPureAIJ = $chJ.SeriesCollection().NewSeries()
    $sPureAIJ.ChartType = 76; $sPureAIJ.AxisGroup = 1
    try { $sPureAIJ.Format.Fill.ForeColor.RGB = $cgdTealMid } catch {}
    try { $sPureAIJ.Format.Fill.Transparency  = 0.5 } catch {}
    try { $sPureAIJ.Format.Line.Visible       = $false } catch {}
    $sPureAIJ.Values  = $wsJ.Range("S3:S${lastRJ}")
    $sPureAIJ.XValues = $wsJ.Range("A3:A${lastRJ}")
}
try {
    $chJ.SeriesCollection(2).Name  = "Annex II"
    $chJ.SeriesCollection(13).Name = "Annex I"
} catch {}
try {
    try { $chJ.Legend.Font.Name = "Sofia Pro" } catch { $chJ.Legend.Font.Name = "Calibri" }
    $chJ.Legend.Font.Size = 12
    $nLeJ = $chJ.Legend.LegendEntries().Count
    for ($le = $nLeJ; $le -ge 1; $le--) {
        if ($le -ne 1 -and $le -ne 3) { try { $chJ.Legend.LegendEntries($le).Delete() } catch {} }
    }
} catch {}

Save-Close $wbJ "$outDir\Figure_12.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
