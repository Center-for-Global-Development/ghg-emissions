<#
.SYNOPSIS
figF.ps1
Generate Figure_8.xlsx -- the colonial-attributed cumulative-shares dumbbell
(1850-2024, mirrors Figure A's structure but with emissions and GMST reassigned
to controlling powers per Carbon Brief's territorial-rule database).

Inputs:
  data/outputs/charts/fig8_dumbbell.csv
  cgd-general/chart-templates/Chart-template-Dumbbell.xlsx

Output:
  data/outputs/charts/Figure_8.xlsx

Pre-req: scripts 11 + 09 must have run so fig8_dumbbell.csv exists.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

# ---------------------------------------------------------------------------
# Colour constants (Excel BGR) -- mirror 10_populate_charts.ps1
# ---------------------------------------------------------------------------
$cgdTeal      = 5983243
$cgdAmber     = 2930175
$cgdTealMid   = 10390042
$cgdAmberDark = 36556
$cgdTealLight = 13679965
$cgdTealMuted = 9866331
$cgdConnLine  = 14407080
$cgdNearBlack = 3355443

$DB_SRC_COLS = @(
    @{ Csv="owid";       Col=3;  Color=$cgdTeal;       Legend="OWID" },
    @{ Csv="primap";     Col=4;  Color=$cgdAmber;      Legend="PRIMAP" },
    @{ Csv="gcp_fossil"; Col=5;  Color=$cgdTealMid;    Legend="GCP" },
    @{ Csv="gcp_blue";   Col=6;  Color=$cgdTealMid;    Legend=$null },
    @{ Csv="gcp_oscar";  Col=7;  Color=$cgdTealMid;    Legend=$null },
    @{ Csv="gcp_luce";   Col=8;  Color=$cgdTealMid;    Legend=$null },
    @{ Csv="edgar";      Col=9;  Color=$cgdAmberDark;  Legend="EDGAR" },
    @{ Csv="cw";         Col=10; Color=$cgdTealLight;  Legend="Climate Watch" }
)
$DB_GMST_COL = 11   # K

. "$PSScriptRoot\_chart_helpers.ps1"

$endash = [char]0x2013
$CHART_TITLE = "Cumulative emissions shares with colonial attribution (1850${endash}2024)"

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

Write-Host "`nFig J: Colonial-attributed dumbbell 1850-2024 ..."

$dCsv  = Load-CSV "fig8_dumbbell.csv"
$nRows = @($dCsv).Count

$wb2 = Open-Template "Chart-template-Dumbbell.xlsx"
$ws2 = $wb2.Sheets("Data")
try { $ws2.Cells.UnMerge() } catch {}
$ws2.Range("A3:S200").ClearContents()

$startRow = 3
$endRow2  = $startRow + $nRows - 1

# --- Write data ---
$r = $startRow
foreach ($row in $dCsv) {
    if ($row.y_pos -ne "" -and $row.y_pos -ne "NaN") {
        $ws2.Cells($r,1).Value2 = [int]$row.y_pos
    }
    if ($row.label -ne "") { $ws2.Cells($r,2).Value2 = [string]$row.label }

    foreach ($def in $DB_SRC_COLS) {
        $v = $row.($def.Csv)
        if ([string]::IsNullOrWhiteSpace($v) -or $v -eq "NaN") {
            $ws2.Cells($r, $def.Col).Formula = "=NA()"
        } else {
            try { $ws2.Cells($r, $def.Col).Value2 = [double]$v } catch {}
        }
    }

    $v = $row.gmst_ref
    if ([string]::IsNullOrWhiteSpace($v) -or $v -eq "NaN") {
        $ws2.Cells($r, $DB_GMST_COL).Formula = "=NA()"
    } else {
        try { $ws2.Cells($r, $DB_GMST_COL).Value2 = [double]$v } catch {}
    }
    $r++
}

# --- Connector formulas L-S ---
for ($r2 = $startRow; $r2 -le $endRow2; $r2++) {
    $ws2.Cells($r2, 12).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,NA(),AGGREGATE(5,6,C${r2}:J${r2}))"
    $ws2.Cells($r2, 13).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,0,AGGREGATE(4,6,C${r2}:J${r2})-AGGREGATE(5,6,C${r2}:J${r2}))"
    $ws2.Cells($r2, 14).Value2  = 0
    $ws2.Cells($r2, 15).Formula = "=IF(ISNA(L${r2}),NA(),L${r2}+M${r2})"
    $ws2.Cells($r2, 16).Formula = "=IF(ISNA(L${r2}),NA(),L${r2})"
    $ws2.Cells($r2, 17).Formula = "=IF(ISNA(L${r2}),NA(),O${r2})"
    $ws2.Cells($r2, 18).Formula = "=IF(ISNA(L${r2}),NA(),A${r2})"
    $ws2.Cells($r2, 19).Formula = "=R${r2}"
}

# --- Header row 2 ---
$ws2.Cells(2, 3).Value2  = "OWID"
$ws2.Cells(2, 4).Value2  = "PRIMAP"
$ws2.Cells(2, 5).Value2  = "GCP"
$ws2.Cells(2, 6).Value2  = "GCP"
$ws2.Cells(2, 7).Value2  = "GCP"
$ws2.Cells(2, 8).Value2  = "GCP"
$ws2.Cells(2, 9).Value2  = "EDGAR"
$ws2.Cells(2, 10).Value2 = "Climate Watch"
$ws2.Cells(2, 11).Value2 = "GMST"

# --- Build chart ---
$chart2 = $wb2.Sheets("Chart").ChartObjects(1).Chart
while ($chart2.SeriesCollection().Count -gt 0) {
    $chart2.SeriesCollection(1).Delete() | Out-Null
}
$rY = $ws2.Range("A${startRow}:A${endRow2}")

foreach ($def in $DB_SRC_COLS) {
    $c  = [char](64 + $def.Col)
    $s2 = $chart2.SeriesCollection().NewSeries()
    $s2.Name      = "=Data!`$${c}`$2"
    $s2.XValues   = $ws2.Range("${c}${startRow}:${c}${endRow2}")
    $s2.Values    = $rY
    $s2.ChartType = -4169
    $s2.MarkerStyle           = 8
    $s2.MarkerSize            = 9
    $s2.MarkerForegroundColor = $def.Color
    $s2.MarkerBackgroundColor = $def.Color
    $s2.HasDataLabels = $false
}

# GMST reference (square marker, no fill, muted teal)
$sRef = $chart2.SeriesCollection().NewSeries()
$sRef.Name      = "=Data!`$K`$2"
$sRef.XValues   = $ws2.Range("K${startRow}:K${endRow2}")
$sRef.Values    = $rY
$sRef.ChartType = -4169
$sRef.MarkerStyle           = 1
$sRef.MarkerSize            = 11
$sRef.MarkerForegroundColor = $cgdTealMuted
try { $sRef.MarkerBackgroundColorIndex = -4142 } catch {}

# _Labels series
$sLbl = $chart2.SeriesCollection().NewSeries()
$sLbl.Name      = "_Labels"
$sLbl.XValues   = $ws2.Range("N${startRow}:N${endRow2}")
$sLbl.Values    = $rY
$sLbl.ChartType = -4169
$sLbl.MarkerStyle = -4142
$sLbl.ApplyDataLabels()
$dlAll = $sLbl.DataLabels()
$dlAll.ShowSeriesName   = $false
$dlAll.ShowCategoryName = $false
try { $dlAll.ShowBubbleSize = $false } catch {}
$dlAll.ShowLegendKey    = $false
$dlAll.Position         = -4131
try { $dlAll.Font.Name = "Sofia Pro" } catch { $dlAll.Font.Name = "Calibri" }
$dlAll.Font.Size  = 11
$dlAll.Font.Color = $cgdNearBlack

$ii = 1
foreach ($row in $dCsv) {
    $lblText = if ($row.label) { [string]$row.label } else { "" }
    try {
        $dl = $sLbl.DataLabels($ii)
        if ($lblText -eq "") { $dl.Formula = '=""' }
        else { $dl.Text = $lblText; Set-SubscriptCO2 $dl }
    } catch {}
    $ii++
}

# Connector series (one per row)
for ($ri = $startRow; $ri -le $endRow2; $ri++) {
    $s = $chart2.SeriesCollection().NewSeries()
    $s.Name      = "_Conn"
    $s.XValues   = $ws2.Range("P${ri}:Q${ri}")
    $s.Values    = $ws2.Range("R${ri}:S${ri}")
    $s.ChartType = 75
    $s.MarkerStyle = -4142
    $s.Format.Line.ForeColor.RGB = $cgdConnLine
    $s.Format.Line.Weight        = 2.25
}

# Legend: keep only data series that actually have values in the CSV
# (EDGAR/CW are all-blank for the colonial 1850 baseline, so they drop out).
$present = @{}
foreach ($def in $DB_SRC_COLS) {
    $has = $false
    foreach ($row in $dCsv) {
        $v = $row.($def.Csv)
        if (-not [string]::IsNullOrWhiteSpace($v) -and $v -ne "NaN") { $has = $true; break }
    }
    $present[$def.Csv] = $has
}
$keepIndices = @()
for ($si = 0; $si -lt $DB_SRC_COLS.Count; $si++) {
    $def = $DB_SRC_COLS[$si]
    if ($null -eq $def.Legend) { continue }
    $has = $present[$def.Csv]
    if ($def.Csv -eq "gcp_fossil") {
        $has = $present["gcp_fossil"] -or $present["gcp_blue"] -or $present["gcp_oscar"] -or $present["gcp_luce"]
    }
    if ($has) { $keepIndices += ($si + 1) }
}
$keepIndices += 9   # GMST always present
$chart2.HasLegend       = $true
$chart2.Legend.Position = -4160
try { $chart2.Legend.Font.Name = "Sofia Pro" } catch { $chart2.Legend.Font.Name = "Calibri" }
try { $chart2.Legend.Font.Size = 12 } catch {}
$nLeg = $chart2.Legend.LegendEntries().Count
for ($li = $nLeg; $li -ge 1; $li--) {
    if ($li -notin $keepIndices) {
        try { $chart2.Legend.LegendEntries($li).Delete() } catch {}
    }
}

# --- Axes ---
$xAx = $chart2.Axes(1)
$xAx.MinimumScale = 0; $xAx.MaximumScale = 100; $xAx.MajorUnit = 10
$xAx.HasTitle = $true
$xAx.AxisTitle.Text = "Share of cumulative emissions (%) or of contribution to rise in GMST"
try { $xAx.AxisTitle.Font.Name = "Sofia Pro" } catch { $xAx.AxisTitle.Font.Name = "Calibri" }
$xAx.AxisTitle.Font.Size = 12
try { $xAx.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx.TickLabels.Font.Name = "Calibri" }
try { $xAx.TickLabels.Font.Size = 12 } catch {}

$yAx = $chart2.Axes(2)
$yAx.MinimumScale = 0; $yAx.MaximumScale = $nRows + 1; $yAx.MajorUnit = 1
$yAx.TickLabels.NumberFormat = ";;;"
$yAx.MajorTickMark = -4142
$yAx.MinorTickMark = -4142
$yAx.HasMajorGridlines = $false
try { $yAx.Format.Line.Visible = $false } catch {}

# --- Title (data sheet only; no title on chart) + layout ---
Set-DataTitle $ws2 $chart2 $CHART_TITLE
try {
    $chart2.PlotArea.InsideLeft   = 220
    $chart2.PlotArea.InsideTop    = 60
    $chart2.PlotArea.InsideWidth  = 580
    $chart2.PlotArea.InsideHeight = 420
} catch {}
Format-DataRange $ws2 $startRow $endRow2 3 11 "0"
Format-Sheet $ws2 $startRow $endRow2 19

Save-Close $wb2 "$outDir\Figure_8.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
