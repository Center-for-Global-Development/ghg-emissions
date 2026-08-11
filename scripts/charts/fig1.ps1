<#
.SYNOPSIS
fig1.ps1
Generate Figure_1.xlsx -- PiePair: GHG composition by gas and sector.

Input:
  data/outputs/charts/fig1_pie.csv
  cgd-general/chart-templates/Chart-template-PiePair.xlsx

Output:
  data/outputs/charts/Figure_1.xlsx

Pre-req: script 09 must have run to produce fig1_pie.csv.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

# CGD palette — applied in rank order (largest slice first)
$SLICE_COLORS = @(5983243, 2930175, 10390042, 36556, 13679965)
# #0B4C5B, #FFB52C, #1A8A9E, #CC8E00, #5DBDD0

. "$PSScriptRoot\_chart_helpers.ps1"

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

Write-Host "`nFigure 1: PiePair ..."
$f1  = Load-CSV "fig1_pie.csv"
$wb  = Open-Template "Chart-template-PiePair.xlsx"
$ws  = $wb.Sheets("Data")

# Sort each breakdown by share descending so slices render largest-first
$gas = @($f1 | Where-Object { $_.breakdown -eq "gas" } |
    Sort-Object { [double]$_.share_pct } -Descending)
$sec = @($f1 | Where-Object { $_.breakdown -eq "sector" } |
    Sort-Object { [double]$_.share_pct } -Descending)

$r = 3
foreach ($row in $gas) {
    $ws.Cells($r,1).Value2 = $row.category
    Set-Value $ws.Cells($r,2) $row.share_pct
    $r++
}
while ($r -le 7) { $ws.Cells($r,1).ClearContents(); $ws.Cells($r,2).ClearContents(); $r++ }

$r = 10
foreach ($row in $sec) {
    $ws.Cells($r,1).Value2 = $row.category
    Set-Value $ws.Cells($r,2) $row.share_pct
    $r++
}
while ($r -le 15) { $ws.Cells($r,1).ClearContents(); $ws.Cells($r,2).ClearContents(); $r++ }

$wsChart1 = $wb.Sheets("Chart")
$gasEnd = 2 + $gas.Count
$secEnd = 9 + $sec.Count
$co1 = $wsChart1.ChartObjects(1); $co2 = $wsChart1.ChartObjects(2)
$co1.Chart.SeriesCollection(1).Values  = $ws.Range("B3:B${gasEnd}")
$co1.Chart.SeriesCollection(1).XValues = $ws.Range("A3:A${gasEnd}")
$co2.Chart.SeriesCollection(1).Values  = $ws.Range("B10:B${secEnd}")
$co2.Chart.SeriesCollection(1).XValues = $ws.Range("A10:A${secEnd}")

# Apply ranked colours to each pie slice (Points are 1-indexed)
for ($i = 0; $i -lt $gas.Count; $i++) {
    $pt = $co1.Chart.SeriesCollection(1).Points($i + 1)
    $pt.Interior.Color = $SLICE_COLORS[$i]
    try { $pt.Border.LineStyle = -4142 } catch {}
}
for ($i = 0; $i -lt $sec.Count; $i++) {
    $pt = $co2.Chart.SeriesCollection(1).Points($i + 1)
    $pt.Interior.Color = $SLICE_COLORS[$i]
    try { $pt.Border.LineStyle = -4142 } catch {}
}

# Title kept in Data!A1 only; no title on the charts themselves
$ws.Cells(1,1).Value2 = "Composition of GHG emissions by gas and sector"
try { $ws.Range("A1:B1").Merge() | Out-Null } catch {}
$co1.Chart.HasTitle = $false
$co2.Chart.HasTitle = $false
foreach ($coX in @($co1, $co2)) {
    try {
        try { $coX.Chart.Legend.Font.Name = "Sofia Pro" } catch { $coX.Chart.Legend.Font.Name = "Calibri" }
        $coX.Chart.Legend.Font.Size = 12
    } catch {}
}
Format-DataRange $ws 3 $gasEnd 2 2 "0.0"
Format-DataRange $ws 10 $secEnd 2 2 "0.0"

Save-Close $wb "$outDir\Figure_1.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
