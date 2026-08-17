<#
.SYNOPSIS
_chart_dumbbell.ps1 - Dumbbell helpers (Figures 3, 4, 8).

Dot-source from any chart script that uses the Dumbbell template:
    . "$PSScriptRoot\_chart_dumbbell.ps1"

Provides:
  $DB_SRC_COLS, $DB_GMST_COL  - source column metadata for the static dumbbell
  Build-Dumbbell                - static cumulative-shares dumbbell

Requires the caller to have set the CGD colour constants ($cgdTeal, $cgdAmber,
$cgdTealMid, $cgdAmberDark, $cgdTealLight, $cgdTealMuted, $cgdConnLine,
$cgdNearBlack) before dot-sourcing, and Set-Value, Load-CSV, Open-Template,
Set-ChartTitle, Format-DataRange, Format-Sheet, Save-Close, Import-CsvToTab,
Add-MeasureTab from _chart_helpers.ps1.

Caller must also have $outDir, $TestPrefix, $MEASURES, $MLABEL, $CHART_TITLES
in scope when invoking these functions.
#>

# Source column definitions: CSV field -> Excel column index, colour, legend label
# GCP variants share cgdTealMid; only the first ("GCP") shows in legend.
$DB_SRC_COLS = @(
    @{ Csv="owid";       Col=3;  Color=$cgdTeal;      Legend="OWID" },
    @{ Csv="primap";     Col=4;  Color=$cgdAmber;      Legend="PRIMAP" },
    @{ Csv="gcp_fossil"; Col=5;  Color=$cgdTealMid;    Legend="GCP" },
    @{ Csv="gcp_blue";   Col=6;  Color=$cgdTealMid;    Legend=$null },
    @{ Csv="gcp_oscar";  Col=7;  Color=$cgdTealMid;    Legend=$null },
    @{ Csv="gcp_luce";   Col=8;  Color=$cgdTealMid;    Legend=$null },
    @{ Csv="edgar";      Col=9;  Color=$cgdAmberDark;  Legend="EDGAR" },
    @{ Csv="cw";         Col=10; Color=$cgdTealLight;  Legend="Climate Watch" }
)
$DB_GMST_COL  = 11   # K

function Build-Dumbbell([string]$csvFile, [string]$outFile, [string]$title) {
    $dCsv = Load-CSV $csvFile
    $nRows = @($dCsv).Count     # 11 rows (8 data + 3 spacers)

    $wb2 = Open-Template "Chart-template-Dumbbell.xlsx"
    $ws2 = $wb2.Sheets("Data")
    try { $ws2.Cells.UnMerge() } catch {}
    $ws2.Range("A3:S200").ClearContents()

    $startRow = 3
    $endRow2  = $startRow + $nRows - 1

    # --- Write data ---
    $r = $startRow
    foreach ($row in $dCsv) {
        # Y_pos: guard blank spacer rows
        if ($row.y_pos -ne "" -and $row.y_pos -ne "NaN") {
            $ws2.Cells($r,1).Value2 = [int]$row.y_pos
        }
        if ($row.label -ne "") { $ws2.Cells($r,2).Value2 = [string]$row.label }

        # Source columns C-J: write NA() for blank so XY scatter treats as gap, not x=0
        foreach ($def in $DB_SRC_COLS) {
            $v = $row.($def.Csv)
            if ([string]::IsNullOrWhiteSpace($v) -or $v -eq "NaN") { $ws2.Cells($r,$def.Col).Formula = "=NA()" }
            else { try { $ws2.Cells($r,$def.Col).Value2 = [double]$v } catch {} }
        }
        # GMST reference (K)
        $v = $row.gmst_ref
        if ([string]::IsNullOrWhiteSpace($v) -or $v -eq "NaN") { $ws2.Cells($r,$DB_GMST_COL).Formula = "=NA()" }
        else { try { $ws2.Cells($r,$DB_GMST_COL).Value2 = [double]$v } catch {} }
        $r++
    }

    # --- Connector formulas L-S (spanning source cols C:J, excluding GMST) ---
    # AGGREGATE(5/4, 6, range) = MIN/MAX ignoring error values (incl. #N/A from blank sources)
    for ($r2 = $startRow; $r2 -le $endRow2; $r2++) {
        $ws2.Cells($r2, 12).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,NA(),AGGREGATE(5,6,C${r2}:J${r2}))"                          # L: Conn_start
        $ws2.Cells($r2, 13).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,0,AGGREGATE(4,6,C${r2}:J${r2})-AGGREGATE(5,6,C${r2}:J${r2}))" # M: Conn_length
        $ws2.Cells($r2, 14).Value2  = 0                                                                         # N: Label_X (always 0)
        $ws2.Cells($r2, 15).Formula = "=IF(ISNA(L${r2}),NA(),L${r2}+M${r2})"                                   # O: Conn_end
        $ws2.Cells($r2, 16).Formula = "=IF(ISNA(L${r2}),NA(),L${r2})"                                           # P: ConnX_1
        $ws2.Cells($r2, 17).Formula = "=IF(ISNA(L${r2}),NA(),O${r2})"                                           # Q: ConnX_2
        $ws2.Cells($r2, 18).Formula = "=IF(ISNA(L${r2}),NA(),A${r2})"                                           # R: ConnY_1
        $ws2.Cells($r2, 19).Formula = "=R${r2}"                                                                  # S: ConnY_2
    }

    # --- Header row 2: series name labels ---
    $ws2.Cells(2, 3).Value2  = "OWID"
    $ws2.Cells(2, 4).Value2  = "PRIMAP"
    $ws2.Cells(2, 5).Value2  = "GCP"
    $ws2.Cells(2, 6).Value2  = "GCP"       # hidden from legend
    $ws2.Cells(2, 7).Value2  = "GCP"       # hidden from legend
    $ws2.Cells(2, 8).Value2  = "GCP"       # hidden from legend
    $ws2.Cells(2, 9).Value2  = "EDGAR"
    $ws2.Cells(2, 10).Value2 = "Climate Watch"
    $ws2.Cells(2, 11).Value2 = "GMST"

    # --- Build chart from scratch ---
    $chart2 = $wb2.Sheets("Chart").ChartObjects(1).Chart
    while ($chart2.SeriesCollection().Count -gt 0) {
        $chart2.SeriesCollection(1).Delete() | Out-Null
    }
    $rY = $ws2.Range("A${startRow}:A${endRow2}")

    # Source dot series (C-J): filled circles, no data labels
    foreach ($def in $DB_SRC_COLS) {
        $c  = [char](64 + $def.Col)      # column letter
        $s2 = $chart2.SeriesCollection().NewSeries()
        $s2.Name      = "=Data!`$${c}`$2"
        $s2.XValues   = $ws2.Range("${c}${startRow}:${c}${endRow2}")
        $s2.Values    = $rY
        $s2.ChartType = -4169             # xlXYScatter
        $s2.MarkerStyle           = 8     # xlMarkerStyleCircle
        $s2.MarkerSize            = 9
        $s2.MarkerForegroundColor = $def.Color
        $s2.MarkerBackgroundColor = $def.Color
        $s2.HasDataLabels = $false        # ensure no stray data labels
    }

    # GMST reference series (K): square marker, no fill, muted teal
    $sRef = $chart2.SeriesCollection().NewSeries()
    $sRef.Name      = "=Data!`$K`$2"
    $sRef.XValues   = $ws2.Range("K${startRow}:K${endRow2}")
    $sRef.Values    = $rY
    $sRef.ChartType = -4169
    $sRef.MarkerStyle           = 1      # xlMarkerStyleSquare
    $sRef.MarkerSize            = 11
    $sRef.MarkerForegroundColor = $cgdTealMuted
    try { $sRef.MarkerBackgroundColorIndex = -4142 } catch {}   # no fill

    # _Labels series (invisible, at X=0, carries data labels)
    $sLbl = $chart2.SeriesCollection().NewSeries()
    $sLbl.Name      = "_Labels"
    $sLbl.XValues   = $ws2.Range("N${startRow}:N${endRow2}")
    $sLbl.Values    = $rY
    $sLbl.ChartType = -4169
    $sLbl.MarkerStyle = -4142             # xlMarkerStyleNone
    $sLbl.ApplyDataLabels()
    $dlAll = $sLbl.DataLabels()
    $dlAll.ShowSeriesName   = $false
    $dlAll.ShowCategoryName = $false
    try { $dlAll.ShowBubbleSize = $false } catch {}
    $dlAll.ShowLegendKey    = $false
    $dlAll.Position         = -4131       # xlLabelPositionLeft
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

    # Connector series: one per row with >=2 source values
    for ($ri = $startRow; $ri -le $endRow2; $ri++) {
        $s = $chart2.SeriesCollection().NewSeries()
        $s.Name      = "_Conn"
        $s.XValues   = $ws2.Range("P${ri}:Q${ri}")
        $s.Values    = $ws2.Range("R${ri}:S${ri}")
        $s.ChartType = 75                 # xlXYScatterLinesNoMarkers
        $s.MarkerStyle = -4142
        $s.Format.Line.ForeColor.RGB = $cgdConnLine
        $s.Format.Line.Weight        = 2.25
    }

    # --- Legend: keep only data series that actually have values in this CSV ---
    # Series order: 1=OWID, 2=PRIMAP, 3=GCP_fossil, 4=GCP_BLUE, 5=GCP_OSCAR,
    #   6=GCP_LUCE, 7=EDGAR, 8=CW, 9=GMST, 10=_Labels, 11..N=_Conn
    # Duplicate GCP variants (4,5,6) never get a legend entry; the first GCP (3)
    # shows if ANY GCP variant has data. Sources absent from the CSV (all-blank
    # columns, e.g. EDGAR/CW in Figs 3/8) are dropped from the legend.
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
    $chart2.HasLegend = $true
    $chart2.Legend.Position = -4160       # xlLegendPositionTop
    try { $chart2.Legend.Font.Name = "Sofia Pro" } catch { $chart2.Legend.Font.Name = "Calibri" }
    try { $chart2.Legend.Font.Size = 12 } catch {}
    $nLeg = $chart2.Legend.LegendEntries().Count
    for ($li = $nLeg; $li -ge 1; $li--) {
        if ($li -notin $keepIndices) {
            try { $chart2.Legend.LegendEntries($li).Delete() } catch {}
        }
    }

    # --- Axis formatting ---
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
    $yAx.TickLabels.NumberFormat = ";;;"       # hide numeric y-axis labels
    $yAx.MajorTickMark = -4142                 # xlNone
    $yAx.MinorTickMark = -4142
    $yAx.HasMajorGridlines = $false
    try { $yAx.Format.Line.Visible = $false } catch {}

    # --- Title (data sheet only; no title on the chart itself) + chart layout ---
    Set-DataTitle $ws2 $chart2 $title
    # Shift plot area down and right to avoid overlap with legend and axis labels
    try {
        $chart2.PlotArea.InsideLeft   = 220   # right to make room for category labels
        $chart2.PlotArea.InsideTop    = 60    # down to make room for legend
        $chart2.PlotArea.InsideWidth  = 580
        $chart2.PlotArea.InsideHeight = 420
    } catch {}
    Format-DataRange $ws2 $startRow $endRow2 3 11 "0"       # source % cols C-K (no decimals)
    Format-Sheet $ws2 $startRow $endRow2 19                 # A-S

    Save-Close $wb2 "$outDir\${TestPrefix}$outFile"
}
