<#
.SYNOPSIS
figA_dynamic.ps1
Build Figure_A_dynamic.xlsx -- the dynamic-date emissions+GMST dumbbell.
GMST sourced from dumbbell_gmst_asis.csv (country-level aggregation via our
Annex flags; same UNFCCC-aligned definition as Figure A static).

Pre-req: scripts 09 must have run so dumbbell_*.csv files exist.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

# ---------------------------------------------------------------------------
# Colour constants + source col defs (mirror 10_populate_charts.ps1)
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

. "$PSScriptRoot\_chart_helpers.ps1"

$endash = [char]0x2013
$em     = [char]0x2014

# ---------------------------------------------------------------------------
# Single-tab SUMIFS / VLOOKUP formula builders (no toggle; matches v2 logic)
# ---------------------------------------------------------------------------
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

function Make-GmstFormula([int]$grpColIdx, [int]$worldColIdx) {
    $p1 = 'Chart!$V$1'; $p2 = 'Chart!$V$2'
    $nc = $worldColIdx + 3
    $rng = "gmst!`$A:`$" + [char](64 + $nc)
    $endV = "VLOOKUP(" + $p2 + "," + $rng + "," + $grpColIdx + ",0)"
    $wldE = "VLOOKUP(" + $p2 + "," + $rng + "," + $worldColIdx + ",0)"
    $preV = "IF(" + $p1 + "<=1851,0,VLOOKUP(" + $p1 + "-1," + $rng + "," + $grpColIdx + ",0))"
    $wldP = "IF(" + $p1 + "<=1851,0,VLOOKUP(" + $p1 + "-1," + $rng + "," + $worldColIdx + ",0))"
    return "=IFERROR((" + $endV + "-" + $preV + ")/(" + $wldE + "-" + $wldP + ")*100,NA())"
}

# ---------------------------------------------------------------------------
Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

Write-Host "`nFigA_dynamic_v3: dynamic dumbbell with country-level GMST ..."

# Row metadata (11 rows: 8 data + 3 spacers)
$ROW_DEFS = @(
    @{ Msr="co2_excLUC"; Grp="annex_1"; Lbl="CO2 excl. LULUCF $em Annex I"      }
    @{ Msr="co2_excLUC"; Grp="annex_2"; Lbl="CO2 excl. LULUCF $em Annex II"     }
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
$nRows = $ROW_DEFS.Count

$MSR_GMST = @{
    "co2_excLUC" = @{ A2=2;  A1=3;  W=4  }
    "co2_incLUC" = @{ A2=5;  A1=6;  W=7  }
    "ghg_excLUC" = @{ A2=8;  A1=9;  W=10 }
    "ghg_incLUC" = @{ A2=11; A1=12; W=13 }
}

$SRC_LABELS    = @("OWID","PRIMAP","GCP fossil","GCP BLUE","GCP OSCAR","GCP LUCE","EDGAR","Climate Watch")
$SRC_MIN_YEARS = @(   0,       0,           0,         0,          0,          0,    1970,          1990)
$RAW_TABS = @{ "co2_excLUC"="raw_co2_excLUC"; "co2_incLUC"="raw_co2_incLUC";
               "ghg_excLUC"="raw_ghg_excLUC"; "ghg_incLUC"="raw_ghg_incLUC" }

$wb = Open-Template "Chart-template-Dumbbell.xlsx"

# Param cells V1/V2 + labels
$chWs = $wb.Sheets("Chart")
$chWs.Cells(1,21).Value2 = "Start year"; $chWs.Cells(2,21).Value2 = "End year"
$chWs.Cells(1,21).Font.Bold = $true; $chWs.Cells(2,21).Font.Bold = $true
try { $chWs.Range("U1:U2").Font.Name = "Sofia Pro" } catch { $chWs.Range("U1:U2").Font.Name = "Calibri" }
$chWs.Cells(1,22).Value2 = 1850; $chWs.Cells(2,22).Value2 = 2024
try {
    $chWs.Range("V1:V2").Interior.Color = 16777164
    $chWs.Range("V1:V2").Font.Bold      = $true
} catch {}

# Import 4 raw tabs + GMST tab (NOTE: gmst_asis, not Jones GLOBAL/ANNEXII)
foreach ($msr in @("co2_excLUC","co2_incLUC","ghg_excLUC","ghg_incLUC")) {
    Import-CsvToTab $wb $RAW_TABS[$msr] "$csvDir\dumbbell_raw_${msr}.csv"
}
Import-CsvToTab $wb "gmst" "$csvDir\dumbbell_gmst_asis.csv"

# Data tab layout
$ws = $wb.Sheets("Data")
try { $ws.Cells.UnMerge() } catch {}
$ws.Range("A3:S200").ClearContents()
$startRow = 3
$endRow   = $startRow + $nRows - 1

for ($si = 0; $si -lt $SRC_LABELS.Count; $si++) {
    $ws.Cells(2, $si + 3).Value2 = $SRC_LABELS[$si]
}
$ws.Cells(2, 11).Value2 = "GMST"

$yPos = 1
for ($ri = 0; $ri -lt $ROW_DEFS.Count; $ri++) {
    $dr  = $startRow + $ri
    $def = $ROW_DEFS[$ri]
    if ($null -eq $def) { $yPos++; continue }
    Set-Value $ws.Cells($dr, 1) $yPos
    $ws.Cells($dr, 2).Value2 = $def.Lbl
    $rawTab = $RAW_TABS[$def.Msr]
    $isA2   = ($def.Grp -eq "annex_2")
    for ($si = 0; $si -lt $SRC_LABELS.Count; $si++) {
        $a2L  = [char](66 + 3*$si)
        $a1L  = [char](67 + 3*$si)
        $na1L = [char](68 + 3*$si)
        $numL = if ($isA2) { [string]$a2L } else { [string]$a1L }
        $sf   = Make-ShareFormula $rawTab $numL ([string]$a1L) ([string]$na1L)
        $minYr = $SRC_MIN_YEARS[$si]
        if ($minYr -gt 0) {
            $ws.Cells($dr, $si + 3).Formula = "=IF(Chart!`$V`$1<$minYr,NA()," + $sf.Substring(1) + ")"
        } else {
            $ws.Cells($dr, $si + 3).Formula = $sf
        }
    }
    $gc   = $MSR_GMST[$def.Msr]
    $gCol = if ($isA2) { $gc.A2 } else { $gc.A1 }
    $ws.Cells($dr, 11).Formula = Make-GmstFormula $gCol $gc.W
    $yPos++
}

# Connector formulas L-S
for ($r2 = $startRow; $r2 -le $endRow; $r2++) {
    $ws.Cells($r2, 12).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,NA(),AGGREGATE(5,6,C${r2}:J${r2}))"
    $ws.Cells($r2, 13).Formula = "=IF(COUNT(C${r2}:J${r2})<=1,0,AGGREGATE(4,6,C${r2}:J${r2})-AGGREGATE(5,6,C${r2}:J${r2}))"
    Set-Value $ws.Cells($r2, 14) 0
    $ws.Cells($r2, 15).Formula = "=IF(ISNA(L${r2}),NA(),L${r2}+M${r2})"
    $ws.Cells($r2, 16).Formula = "=IF(ISNA(L${r2}),NA(),L${r2})"
    $ws.Cells($r2, 17).Formula = "=IF(ISNA(L${r2}),NA(),O${r2})"
    $ws.Cells($r2, 18).Formula = "=IF(ISNA(L${r2}),NA(),A${r2})"
    $ws.Cells($r2, 19).Formula = "=R${r2}"
}

# Chart
$chart = $wb.Sheets("Chart").ChartObjects(1).Chart
while ($chart.SeriesCollection().Count -gt 0) {
    $chart.SeriesCollection(1).Delete() | Out-Null
}
$rY = $ws.Range("A${startRow}:A${endRow}")

foreach ($def in $DB_SRC_COLS) {
    $c = [char](64 + $def.Col)
    $s = $chart.SeriesCollection().NewSeries()
    $s.Name      = "=Data!`$${c}`$2"
    $s.XValues   = $ws.Range("${c}${startRow}:${c}${endRow}")
    $s.Values    = $rY
    $s.ChartType = -4169
    $s.MarkerStyle           = 8
    $s.MarkerSize            = 9
    $s.MarkerForegroundColor = $def.Color
    $s.MarkerBackgroundColor = $def.Color
    $s.HasDataLabels = $false
}

$sRef = $chart.SeriesCollection().NewSeries()
$sRef.Name      = "=Data!`$K`$2"
$sRef.XValues   = $ws.Range("K${startRow}:K${endRow}")
$sRef.Values    = $rY
$sRef.ChartType = -4169
$sRef.MarkerStyle           = 1
$sRef.MarkerSize            = 11
$sRef.MarkerForegroundColor = $cgdTealMuted
try { $sRef.MarkerBackgroundColorIndex = -4142 } catch {}

$sLbl = $chart.SeriesCollection().NewSeries()
$sLbl.Name      = "_Labels"
$sLbl.XValues   = $ws.Range("N${startRow}:N${endRow}")
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
foreach ($def in $ROW_DEFS) {
    $txt = if ($def -ne $null) { $def.Lbl } else { "" }
    try { $sLbl.DataLabels($ii).Formula = "=" + '"' + $txt + '"' } catch {}
    $ii++
}

for ($ri = $startRow; $ri -le $endRow; $ri++) {
    $s = $chart.SeriesCollection().NewSeries()
    $s.Name = "_Conn"
    $s.XValues = $ws.Range("P${ri}:Q${ri}")
    $s.Values  = $ws.Range("R${ri}:S${ri}")
    $s.ChartType = 75
    $s.MarkerStyle = -4142
    $s.Format.Line.ForeColor.RGB = $cgdConnLine
    $s.Format.Line.Weight        = 2.25
}

$chart.HasLegend = $true
$chart.Legend.Position = -4160
try { $chart.Legend.Font.Name = "Sofia Pro" } catch { $chart.Legend.Font.Name = "Calibri" }
try { $chart.Legend.Font.Size = 12 } catch {}
$keepIndices = @(1,2,3,7,8,9)
$nLeg = $chart.Legend.LegendEntries().Count
for ($li = $nLeg; $li -ge 1; $li--) {
    if ($li -notin $keepIndices) {
        try { $chart.Legend.LegendEntries($li).Delete() } catch {}
    }
}

$xAx = $chart.Axes(1)
$xAx.MinimumScale = 0; $xAx.MaximumScale = 100; $xAx.MajorUnit = 10
$xAx.HasTitle = $true
$xAx.AxisTitle.Text = "Share of cumulative emissions (%) or of contribution to rise in GMST"
try { $xAx.AxisTitle.Font.Name = "Sofia Pro" } catch { $xAx.AxisTitle.Font.Name = "Calibri" }
$xAx.AxisTitle.Font.Size = 12
try { $xAx.TickLabels.Font.Name = "Sofia Pro" } catch { $xAx.TickLabels.Font.Name = "Calibri" }
try { $xAx.TickLabels.Font.Size = 12 } catch {}
$yAx = $chart.Axes(2)
$yAx.MinimumScale = 0; $yAx.MaximumScale = $nRows + 1; $yAx.MajorUnit = 1
$yAx.TickLabels.NumberFormat = ";;;"
$yAx.MajorTickMark = -4142
$yAx.MinorTickMark = -4142
$yAx.HasMajorGridlines = $false
try { $yAx.Format.Line.Visible = $false } catch {}

# Dynamic title
try { $ws.Range("A1:S1").Merge() | Out-Null } catch {}
$ws.Cells(1,1).Formula = '="Emissions shares and contribution to increase in GMST ("&TEXT(Chart!$V$1,"0")&"' + $endash + '"&TEXT(Chart!$V$2,"0")&")"'
$chart.HasTitle = $true
$chart.ChartTitle.Formula = "='Data'!`$A`$1"
try {
    $chart.PlotArea.InsideLeft   = 220
    $chart.PlotArea.InsideTop    = 60
    $chart.PlotArea.InsideWidth  = 580
    $chart.PlotArea.InsideHeight = 420
} catch {}

Format-DataRange $ws $startRow $endRow 3 11 "0.0"
Format-Sheet $ws $startRow $endRow 19

Save-Close $wb "$outDir\Figure_A_dynamic.xlsx"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
