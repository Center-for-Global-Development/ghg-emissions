<#
.SYNOPSIS
figH.ps1 — Cumulative per-capita emissions dumbbell, 1850–2024 baseline (Figure 10 in paper).

Reads fig10_data.csv (built by 09_prepare_chart_data.py).
Three dots per row: Method A (TS), Method B (EN), Method C (PYW) + abs_share triangle.

Input:
  data/outputs/charts/fig10_data.csv
  templates/Chart-template-Dumbbell.xlsx

Output:
  data/outputs/charts/Figure_10.xlsx

For the 1990 baseline (Figure I), see figI.ps1.
Pre-req: scripts 01, 09 must have run to produce the input CSV.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

. "$PSScriptRoot\_chart_helpers.ps1"

$cgdTeal      = 5983243
$cgdAmber     = 2930175
$cgdTealMid   = 10390042
$cgdConnLine  = 14407080
$cgdNearBlack = 3355443

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

# ---------------------------------------------------------------------------
# Build-FigH: dumbbell with 3 method dots + 1 absolute-share triangle per row.
#
# CSV columns: y_pos, label, method_a (TS), method_b (EN), method_c (PYW), abs_share
# Data sheet layout:
#   A  = y_pos          B  = label
#   C  = TS             D  = EN             E  = PYW
#   F  = abs_share      (outside connector range; triangle marker, not connected)
#   G  = Conn_start     H  = Conn_length    I  = Label_X (0)
#   J  = Conn_end       K  = ConnX_1        L  = ConnX_2
#   M  = ConnY_1        N  = ConnY_2
# ---------------------------------------------------------------------------
function Build-FigH([string]$csvFile, [string]$outFile, [string]$title) {
    $dCsv = Load-CSV $csvFile
    $nRows = @($dCsv).Count    # 11 rows (8 data + 3 spacers)

    $wb = Open-Template "Chart-template-Dumbbell.xlsx"
    $ws = $wb.Sheets("Data")
    try { $ws.Cells.UnMerge() } catch {}
    $ws.Range("A3:S200").ClearContents()

    $startRow = 3
    $endRow   = $startRow + $nRows - 1

    # --- Write data ---
    $r = $startRow
    foreach ($row in $dCsv) {
        if ($row.y_pos -ne "" -and $row.y_pos -ne "NaN") {
            $ws.Cells($r, 1).Value2 = [int]$row.y_pos
        }
        if ($row.label -ne "") { $ws.Cells($r, 2).Value2 = [string]$row.label }

        foreach ($col in @(
            @{Field="method_a";  Idx=3},
            @{Field="method_b";  Idx=4},
            @{Field="method_c";  Idx=5},
            @{Field="abs_share"; Idx=6}
        )) {
            $v = $row.($col.Field)
            if ([string]::IsNullOrWhiteSpace($v) -or $v -eq "NaN") {
                $ws.Cells($r, $col.Idx).Formula = "=NA()"
            } else {
                try { $ws.Cells($r, $col.Idx).Value2 = [double]$v } catch {}
            }
        }
        $r++
    }

    # --- Connector formulas G-N (span C:E only — abs_share in F is not connected) ---
    for ($r2 = $startRow; $r2 -le $endRow; $r2++) {
        $ws.Cells($r2,  7).Formula = "=IF(COUNT(C${r2}:E${r2})<=1,NA(),AGGREGATE(5,6,C${r2}:E${r2}))"                             # G: Conn_start
        $ws.Cells($r2,  8).Formula = "=IF(COUNT(C${r2}:E${r2})<=1,0,AGGREGATE(4,6,C${r2}:E${r2})-AGGREGATE(5,6,C${r2}:E${r2}))"  # H: Conn_length
        $ws.Cells($r2,  9).Value2  = 0                                                                                              # I: Label_X
        $ws.Cells($r2, 10).Formula = "=IF(ISNA(G${r2}),NA(),G${r2}+H${r2})"                                                        # J: Conn_end
        $ws.Cells($r2, 11).Formula = "=IF(ISNA(G${r2}),NA(),G${r2})"                                                               # K: ConnX_1
        $ws.Cells($r2, 12).Formula = "=IF(ISNA(G${r2}),NA(),J${r2})"                                                               # L: ConnX_2
        $ws.Cells($r2, 13).Formula = "=IF(ISNA(G${r2}),NA(),A${r2})"                                                               # M: ConnY_1
        $ws.Cells($r2, 14).Formula = "=M${r2}"                                                                                      # N: ConnY_2
    }

    # --- Series name headers (row 2) ---
    $ws.Cells(2, 3).Value2 = "Trajectory-summed (TS)"
    $ws.Cells(2, 4).Value2 = "Endpoint-normalised (EN)"
    $ws.Cells(2, 5).Value2 = "Person-year-weighted (PYW)"
    $ws.Cells(2, 6).Value2 = "Cumulative absolute share (OWID)"

    # --- Build chart from scratch ---
    $chart = $wb.Sheets("Chart").ChartObjects(1).Chart
    while ($chart.SeriesCollection().Count -gt 0) {
        $chart.SeriesCollection(1).Delete() | Out-Null
    }
    $rY = $ws.Range("A${startRow}:A${endRow}")

    # Three method dot series: filled circles
    $methodSeries = @(
        @{ Col=3; Color=$cgdTeal;    Name="=Data!`$C`$2" },
        @{ Col=4; Color=$cgdAmber;   Name="=Data!`$D`$2" },
        @{ Col=5; Color=$cgdTealMid; Name="=Data!`$E`$2" }
    )
    foreach ($def in $methodSeries) {
        $c  = [char](64 + $def.Col)
        $s  = $chart.SeriesCollection().NewSeries()
        $s.Name      = $def.Name
        $s.XValues   = $ws.Range("${c}${startRow}:${c}${endRow}")
        $s.Values    = $rY
        $s.ChartType = -4169          # xlXYScatter
        $s.MarkerStyle           = 8  # xlMarkerStyleCircle
        $s.MarkerSize            = 9
        $s.MarkerForegroundColor = $def.Color
        $s.MarkerBackgroundColor = $def.Color
        $s.HasDataLabels         = $false
    }

    # Absolute share series: filled triangle, dark teal — not connected to bar
    $sAbs = $chart.SeriesCollection().NewSeries()
    $sAbs.Name      = "=Data!`$F`$2"
    $sAbs.XValues   = $ws.Range("F${startRow}:F${endRow}")
    $sAbs.Values    = $rY
    $sAbs.ChartType = -4169
    $sAbs.MarkerStyle           = 3   # xlMarkerStyleTriangle
    $sAbs.MarkerSize            = 9
    $sAbs.MarkerForegroundColor = $cgdTeal
    $sAbs.MarkerBackgroundColor = $cgdTeal
    $sAbs.HasDataLabels         = $false

    # _Labels series at X=0, carries row labels
    $sLbl = $chart.SeriesCollection().NewSeries()
    $sLbl.Name      = "_Labels"
    $sLbl.XValues   = $ws.Range("I${startRow}:I${endRow}")
    $sLbl.Values    = $rY
    $sLbl.ChartType = -4169
    $sLbl.MarkerStyle = -4142         # xlMarkerStyleNone
    $sLbl.ApplyDataLabels()
    $dlAll = $sLbl.DataLabels()
    $dlAll.ShowSeriesName   = $false
    $dlAll.ShowCategoryName = $false
    try { $dlAll.ShowBubbleSize = $false } catch {}
    $dlAll.ShowLegendKey = $false
    $dlAll.Position      = -4131      # xlLabelPositionLeft
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

    # Connector series: one per row, spans C:E only (abs_share not included)
    for ($ri = $startRow; $ri -le $endRow; $ri++) {
        $s = $chart.SeriesCollection().NewSeries()
        $s.Name      = "_Conn"
        $s.XValues   = $ws.Range("K${ri}:L${ri}")
        $s.Values    = $ws.Range("M${ri}:N${ri}")
        $s.ChartType = 75             # xlXYScatterLinesNoMarkers
        $s.MarkerStyle = -4142
        $s.Format.Line.ForeColor.RGB = $cgdConnLine
        $s.Format.Line.Weight        = 2.25
    }

    # --- Legend: keep method series (1-3) and abs_share (4) ---
    $chart.HasLegend = $true
    $chart.Legend.Position = -4160    # xlLegendPositionTop
    try { $chart.Legend.Font.Name = "Sofia Pro" } catch { $chart.Legend.Font.Name = "Calibri" }
    try { $chart.Legend.Font.Size = 12 } catch {}
    $nLeg = $chart.Legend.LegendEntries().Count
    for ($li = $nLeg; $li -ge 1; $li--) {
        if ($li -notin @(1, 2, 3, 4)) {
            try { $chart.Legend.LegendEntries($li).Delete() } catch {}
        }
    }

    # --- Axes ---
    $xAx = $chart.Axes(1)
    $xAx.MinimumScale = 0; $xAx.MaximumScale = 60; $xAx.MajorUnit = 10
    $xAx.HasTitle = $true
    $xAx.AxisTitle.Text = "% of world per-capita cumulative emissions"
    try { $xAx.AxisTitle.Font.Name = "Sofia Pro" } catch { $xAx.AxisTitle.Font.Name = "Calibri" }
    $xAx.AxisTitle.Font.Size = 12
    try { $xAx.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx.TickLabels.Font.Name = "Calibri" }
    try { $xAx.TickLabels.Font.Size = 12 } catch {}

    $yAx = $chart.Axes(2)
    $yAx.MinimumScale = 0; $yAx.MaximumScale = $nRows + 1; $yAx.MajorUnit = 1
    $yAx.TickLabels.NumberFormat = ";;;"
    $yAx.MajorTickMark = -4142; $yAx.MinorTickMark = -4142
    $yAx.HasMajorGridlines = $false
    try { $yAx.Format.Line.Visible = $false } catch {}

    # --- Title (data sheet only; no title on chart) + layout ---
    Set-DataTitle $ws $chart $title
    try {
        $chart.PlotArea.InsideLeft   = 240
        $chart.PlotArea.InsideTop    = 60
        $chart.PlotArea.InsideWidth  = 560
        $chart.PlotArea.InsideHeight = 420
    } catch {}
    Format-DataRange $ws $startRow $endRow 3 6 "0.0"   # data cols C-F
    Format-Sheet     $ws $startRow $endRow 14           # A-N

    Save-Close $wb "$outDir\$outFile"
}

try {
    $em = [char]0x2013   # en dash

    Write-Host "`nFigure 10: 1850-2024 baseline ..."
    Build-FigH "fig10_data.csv" "Figure_10.xlsx" `
        "Cumulative per-capita emissions as % of world total (1850${em}2024)"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
