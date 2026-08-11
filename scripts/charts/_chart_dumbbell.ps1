<#
.SYNOPSIS
_chart_dumbbell.ps1 - Dumbbell helpers (Figs A, B, A-dynamic).

Dot-source from any chart script that uses the Dumbbell template:
    . "$PSScriptRoot\_chart_dumbbell.ps1"

Provides:
  $DB_SRC_COLS, $DB_GMST_COL  - source column metadata for the static dumbbell
  Build-Dumbbell                - figA/figB static cumulative-shares dumbbell
  Make-ShareFormula             - SUMIFS share formula builder for the dynamic dumbbell
  Make-GmstFormula              - GMST cumulative formula builder for the dynamic dumbbell
  Build-Dumbbell-v4             - figA-dynamic with user-editable start/end year cells

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

# ===========================================================================
# Helpers for Build-Dumbbell-v4 (dynamic date dumbbell)
# ===========================================================================

# Import-CsvToTab is in _chart_helpers.ps1 (dot-sourced at top).

# Build SUMIFS formula: SUMIFS(numCol)/( SUMIFS(a1Col)+SUMIFS(na1Col) ) * 100
# with year range from Chart sheet cells V1 (start) and V2 (end).
function Make-ShareFormula([string]$tab, [string]$numL, [string]$a1L, [string]$na1L) {
    $q = '"'; $yr = '$A:$A'; $p1 = 'Chart!$V$1'; $p2 = 'Chart!$V$2'
    $mkS = {
        param([string]$col)
        "SUMIFS(" + $tab + "!" + '$' + $col + ":" + '$' + $col + "," +
        $tab + "!" + $yr + "," + $q + ">=" + $q + "&" + $p1 + "," +
        $tab + "!" + $yr + "," + $q + "<=" + $q + "&" + $p2 + ")"
    }
    $fN  = & $mkS $numL; $fD1 = & $mkS $a1L; $fD2 = & $mkS $na1L
    return "=IFERROR(" + $fN + "/(" + $fD1 + "+" + $fD2 + ")*100,NA())"
}

# Build VLOOKUP formula for GMST cumulative share over [start, end] year range.
# dumbbell_gmst.csv: year | a2_co2_excLUC(2) a1(3) w(4) | a2_co2_incLUC(5) a1(6) w(7) | ...
# grpColIdx: VLOOKUP col for numerator; worldColIdx: VLOOKUP col for denominator.
function Make-GmstFormula([int]$grpColIdx, [int]$worldColIdx) {
    $q = '"'; $p1 = 'Chart!$V$1'; $p2 = 'Chart!$V$2'
    $nc = $worldColIdx + 3
    $rng = "gmst!`$A:`$" + [char](64 + $nc)
    $endV = "VLOOKUP(" + $p2 + "," + $rng + "," + $grpColIdx + ",0)"
    $wldE = "VLOOKUP(" + $p2 + "," + $rng + "," + $worldColIdx + ",0)"
    $preV = "IF(" + $p1 + "<=1851,0,VLOOKUP(" + $p1 + "-1," + $rng + "," + $grpColIdx + ",0))"
    $wldP = "IF(" + $p1 + "<=1851,0,VLOOKUP(" + $p1 + "-1," + $rng + "," + $worldColIdx + ",0))"
    return "=IFERROR((" + $endV + "-" + $preV + ")/(" + $wldE + "-" + $wldP + ")*100,NA())"
}

# Build dynamic-date dumbbell workbook.
# User edits Chart!V1 (start year) and Chart!V2 (end year) to recompute all SUMIFS.
function Build-Dumbbell-v4([string]$outFile, [string]$title, [int]$defStart, [int]$defEnd) {
    # Row metadata (11 rows: 8 data + 3 spacers). Order matches figA/B CSV row order.
    # Measure order: co2_excLUC, ghg_excLUC, co2_incLUC, ghg_incLUC (per build_dumbbell in 09).
    # GMST col indices per measure (dumbbell_gmst.csv col order):
    # year | a2_co2_excLUC(2) a1(3) w(4) | a2_co2_incLUC(5) a1(6) w(7)
    #      | a2_ghg_excLUC(8) a1(9) w(10) | a2_ghg_incLUC(11) a1(12) w(13)
    $MSR_GMST_V4 = @{
        "co2_excLUC" = @{A2=2;  A1=3;  W=4}
        "co2_incLUC" = @{A2=5;  A1=6;  W=7}
        "ghg_excLUC" = @{A2=8;  A1=9;  W=10}
        "ghg_incLUC" = @{A2=11; A1=12; W=13}
    }
    $em = [char]0x2014   # em dash; avoid literal in script due to PS5.1 encoding
    $ROW_DEFS = @(
        @{ Msr="co2_excLUC"; Grp="annex_1"; Lbl="CO2 excl. LULUCF $em Annex I"       }
        @{ Msr="co2_excLUC"; Grp="annex_2"; Lbl="CO2 excl. LULUCF $em Annex II"      }
        $null
        @{ Msr="ghg_excLUC"; Grp="annex_1"; Lbl="All GHG excl. LULUCF $em Annex I"  }
        @{ Msr="ghg_excLUC"; Grp="annex_2"; Lbl="All GHG excl. LULUCF $em Annex II" }
        $null
        @{ Msr="co2_incLUC"; Grp="annex_1"; Lbl="CO2 incl. LULUCF $em Annex I"      }
        @{ Msr="co2_incLUC"; Grp="annex_2"; Lbl="CO2 incl. LULUCF $em Annex II"     }
        $null
        @{ Msr="ghg_incLUC"; Grp="annex_1"; Lbl="All GHG incl. LULUCF $em Annex I"  }
        @{ Msr="ghg_incLUC"; Grp="annex_2"; Lbl="All GHG incl. LULUCF $em Annex II" }
    )
    # Source order matches DB_ALL_SOURCES_RAW in script 09.
    # In raw tabs: col A=year, B-D=OWID(a2,a1,na1), E-G=PRIMAP, H-J=GCP_fossil,
    #   K-M=GCP_BLUE, N-P=GCP_OSCAR, Q-S=GCP_LUCE, T-V=EDGAR, W-Y=CW.
    # Data tab source columns: C=OWID, D=PRIMAP, E=GCP_fossil, ..., J=CW (matches DB_SRC_COLS).
    $SRC_LABELS    = @("OWID","PRIMAP","GCP fossil","GCP BLUE","GCP OSCAR","GCP LUCE","EDGAR","Climate Watch")
    # Min start year per source: if Chart!V1 < min, show NA() (source invisible for that range)
    $SRC_MIN_YEARS = @(   0,       0,           0,         0,          0,          0,    1970,          1990)
    $RAW_TABS   = @{ "co2_excLUC"="raw_co2_excLUC"; "co2_incLUC"="raw_co2_incLUC";
                     "ghg_excLUC"="raw_ghg_excLUC"; "ghg_incLUC"="raw_ghg_incLUC" }

    $wb2 = Open-Template "Chart-template-Dumbbell.xlsx"

    # Params on Chart sheet col U-V (editable by user; no separate tab needed)
    $chWs2 = $wb2.Sheets("Chart")
    $chWs2.Cells(1,21).Value2 = "Start year"; $chWs2.Cells(2,21).Value2 = "End year"
    $chWs2.Cells(1,21).Font.Bold = $true; $chWs2.Cells(2,21).Font.Bold = $true
    try { $chWs2.Range("U1:U2").Font.Name = "Sofia Pro" } catch { $chWs2.Range("U1:U2").Font.Name = "Calibri" }
    $chWs2.Cells(1,22).Value2 = $defStart; $chWs2.Cells(2,22).Value2 = $defEnd
    try {
        $chWs2.Cells(1,22).Interior.Color = 16777164
        $chWs2.Cells(2,22).Interior.Color = 16777164
        $chWs2.Cells(1,22).Font.Bold = $true; $chWs2.Cells(2,22).Font.Bold = $true
    } catch {}

    # Add raw data tabs (4 measures) and gmst tab via QueryTable (avoids PS5.1 array marshalling)
    foreach ($msr in @("co2_excLUC","co2_incLUC","ghg_excLUC","ghg_incLUC")) {
        Import-CsvToTab $wb2 $RAW_TABS[$msr] "$csvDir\dumbbell_raw_${msr}.csv"
    }
    # Use dumbbell_gmst_asis.csv (country-level aggregation via our annex flags),
    # NOT dumbbell_gmst.csv (Jones et al.'s pre-aggregated GLOBAL/ANNEXII rows).
    # The country-level path matches static FigA, FigJ, and FigJ_dynamic so all
    # four charts use one consistent UNFCCC-aligned Annex definition (KAZ + MCO
    # in Annex I). See docs/colonial_attribution_method.md for the rationale.
    Import-CsvToTab $wb2 "gmst" "$csvDir\dumbbell_gmst_asis.csv"

    # Build Data tab
    $ws2 = $wb2.Sheets("Data")
    try { $ws2.Cells.UnMerge() } catch {}
    $ws2.Range("A3:S200").ClearContents()
    $nRows    = $ROW_DEFS.Count   # 11
    $startRow = 3
    $endRow2  = $startRow + $nRows - 1

    # Source header labels in row 2 (C=col3 .. J=col10, K=GMST)
    for ($si = 0; $si -lt $SRC_LABELS.Count; $si++) {
        $ws2.Cells(2, $si + 3).Value2 = $SRC_LABELS[$si]
    }
    $ws2.Cells(2, 11).Value2 = "GMST"

    # Write y_pos, labels, SUMIFS formulas, connector placeholders
    $yPos = 1
    for ($ri = 0; $ri -lt $ROW_DEFS.Count; $ri++) {
        $dr  = $startRow + $ri
        $def = $ROW_DEFS[$ri]
        if ($null -eq $def) { $yPos++; continue }   # spacer — leave row blank
        Set-Value $ws2.Cells($dr, 1) $yPos
        $ws2.Cells($dr, 2).Value2 = $def.Lbl
        $rawTab = $RAW_TABS[$def.Msr]
        $isA2   = ($def.Grp -eq "annex_2")
        # Source columns C-J (Data tab cols 3-10)
        for ($si = 0; $si -lt $SRC_LABELS.Count; $si++) {
            $a2L  = [char](66 + 3*$si)   # B, E, H, K, N, Q, T, W
            $a1L  = [char](67 + 3*$si)   # C, F, I, L, O, R, U, X
            $na1L = [char](68 + 3*$si)   # D, G, J, M, P, S, V, Y
            $numL = if ($isA2) { [string]$a2L } else { [string]$a1L }
            $sf   = Make-ShareFormula $rawTab $numL ([string]$a1L) ([string]$na1L)
            $minYr = $SRC_MIN_YEARS[$si]
            if ($minYr -gt 0) {
                # Hide source when start year precedes its coverage
                $ws2.Cells($dr, $si + 3).Formula = "=IF(Chart!`$V`$1<$minYr,NA()," + $sf.Substring(1) + ")"
            } else {
                $ws2.Cells($dr, $si + 3).Formula = $sf
            }
        }
        # GMST column K (col 11) — per-measure gas/component
        $gc   = $MSR_GMST_V4[$def.Msr]
        $gCol = if ($isA2) { $gc.A2 } else { $gc.A1 }
        $ws2.Cells($dr, 11).Formula = Make-GmstFormula $gCol $gc.W
        $yPos++
    }

    # Connector formulas L-S — AGGREGATE ignores #N/A from sources with no data
    for ($r2 = $startRow; $r2 -le $endRow2; $r2++) {
        $ws2.Cells($r2, 12).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,NA(),AGGREGATE(5,6,C${r2}:J${r2}))"
        $ws2.Cells($r2, 13).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,0,AGGREGATE(4,6,C${r2}:J${r2})-AGGREGATE(5,6,C${r2}:J${r2}))"
        Set-Value $ws2.Cells($r2, 14) 0
        $ws2.Cells($r2, 15).Formula = "=IF(ISNA(L${r2}),NA(),L${r2}+M${r2})"
        $ws2.Cells($r2, 16).Formula = "=IF(ISNA(L${r2}),NA(),L${r2})"
        $ws2.Cells($r2, 17).Formula = "=IF(ISNA(L${r2}),NA(),O${r2})"
        $ws2.Cells($r2, 18).Formula = "=IF(ISNA(L${r2}),NA(),A${r2})"
        $ws2.Cells($r2, 19).Formula = "=R${r2}"
    }

    # Build chart (same as Build-Dumbbell; GMST uses × marker instead of hollow circle)
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
        $s2.MarkerStyle           = 8; $s2.MarkerSize = 9
        $s2.MarkerForegroundColor = $def.Color
        $s2.MarkerBackgroundColor = $def.Color
        $s2.HasDataLabels = $false
    }

    # GMST: × marker, no fill, muted teal
    $sRef = $chart2.SeriesCollection().NewSeries()
    $sRef.Name      = "=Data!`$K`$2"
    $sRef.XValues   = $ws2.Range("K${startRow}:K${endRow2}")
    $sRef.Values    = $rY
    $sRef.ChartType = -4169
    $sRef.MarkerStyle           = 1        # xlMarkerStyleSquare
    $sRef.MarkerSize            = 11
    $sRef.MarkerForegroundColor = $cgdTealMuted
    try { $sRef.MarkerBackgroundColorIndex = -4142 } catch {}   # no fill

    # _Labels series
    $sLbl = $chart2.SeriesCollection().NewSeries()
    $sLbl.Name      = "_Labels"
    $sLbl.XValues   = $ws2.Range("N${startRow}:N${endRow2}")
    $sLbl.Values    = $rY
    $sLbl.ChartType = -4169
    $sLbl.MarkerStyle = -4142
    $sLbl.ApplyDataLabels()
    $dlAll = $sLbl.DataLabels()
    $dlAll.ShowSeriesName = $false; $dlAll.ShowCategoryName = $false
    try { $dlAll.ShowBubbleSize = $false } catch {}
    $dlAll.ShowLegendKey = $false
    $dlAll.Position      = -4131   # xlLabelPositionLeft
    try { $dlAll.Font.Name = "Sofia Pro" } catch { $dlAll.Font.Name = "Calibri" }
    $dlAll.Font.Size = 11; $dlAll.Font.Color = $cgdNearBlack
    # Labels: use static text (params change values, not labels)
    $ii = 1
    foreach ($def in $ROW_DEFS) {
        $txt = if ($def -ne $null) { $def.Lbl } else { "" }
        try { $sLbl.DataLabels($ii).Formula = "=" + '"' + $txt + '"' } catch {}
        $ii++
    }

    # Connector series (one per row)
    for ($ri2 = $startRow; $ri2 -le $endRow2; $ri2++) {
        $s = $chart2.SeriesCollection().NewSeries()
        $s.Name = "_Conn"; $s.XValues = $ws2.Range("P${ri2}:Q${ri2}")
        $s.Values = $ws2.Range("R${ri2}:S${ri2}"); $s.ChartType = 75
        $s.MarkerStyle = -4142
        $s.Format.Line.ForeColor.RGB = $cgdConnLine; $s.Format.Line.Weight = 2.25
    }

    # Legend: keep 1,2,3,7,8,9 (OWID, PRIMAP, GCP, EDGAR, CW, GMST)
    # Static keep-list retained here: the dynamic dumbbell's sources appear and
    # disappear with the user-set year range, so all legend entries stay.
    $chart2.HasLegend = $true
    $chart2.Legend.Position = -4160
    try { $chart2.Legend.Font.Name = "Sofia Pro" } catch { $chart2.Legend.Font.Name = "Calibri" }
    try { $chart2.Legend.Font.Size = 12 } catch {}
    $nLeg4 = $chart2.Legend.LegendEntries().Count
    for ($li = $nLeg4; $li -ge 1; $li--) {
        if ($li -notin @(1,2,3,7,8,9)) {
            try { $chart2.Legend.LegendEntries($li).Delete() } catch {}
        }
    }

    # Axes
    $xAx4 = $chart2.Axes(1)
    $xAx4.MinimumScale=0; $xAx4.MaximumScale=100; $xAx4.MajorUnit=10
    $xAx4.HasTitle = $true
    $xAx4.AxisTitle.Text = "Share of cumulative emissions (%) or of contribution to rise in GMST"
    try { $xAx4.AxisTitle.Font.Name = "Sofia Pro" } catch { $xAx4.AxisTitle.Font.Name = "Calibri" }
    $xAx4.AxisTitle.Font.Size = 12
    try { $xAx4.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx4.TickLabels.Font.Name = "Calibri" }
    try { $xAx4.TickLabels.Font.Size = 12 } catch {}
    $yAx4 = $chart2.Axes(2)
    $yAx4.MinimumScale=0; $yAx4.MaximumScale=$nRows+1; $yAx4.MajorUnit=1
    $yAx4.TickLabels.NumberFormat = ";;;"; $yAx4.MajorTickMark = -4142; $yAx4.MinorTickMark = -4142
    $yAx4.HasMajorGridlines = $false
    try { $yAx4.Format.Line.Visible = $false } catch {}

    # Dynamic title: concatenates date range from Chart!V1/V2 so it updates when params change
    try { $ws2.Range("A1:S1").Merge() | Out-Null } catch {}
    $endash = [char]0x2013
    $ws2.Cells(1,1).Formula = '="Emissions shares and contribution to increase in GMST ("&TEXT(Chart!$V$1,"0")&"' + $endash + '"&TEXT(Chart!$V$2,"0")&")"'
    $chart2.HasTitle = $true
    $chart2.ChartTitle.Formula = "='Data'!`$A`$1"
    try {
        $chart2.PlotArea.InsideLeft=220; $chart2.PlotArea.InsideTop=60
        $chart2.PlotArea.InsideWidth=580; $chart2.PlotArea.InsideHeight=420
    } catch {}
    Format-DataRange $ws2 $startRow $endRow2 3 11 "0.0"
    Format-Sheet $ws2 $startRow $endRow2 19

    Save-Close $wb2 "$outDir\${TestPrefix}$outFile"
}
