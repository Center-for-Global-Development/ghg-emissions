<#
.SYNOPSIS
_chart_helpers.ps1 — shared helpers for chart-generation scripts.

Dot-source from any chart script:
    . "$PSScriptRoot\_chart_helpers.ps1"

Requires the caller to have set $csvDir and $TmplDir before dot-sourcing,
and $xl (the Excel COM object) before calling Open-Template.
#>

function Load-CSV([string]$fname) {
    return Import-Csv "$csvDir\$fname"
}

function Get-Num($val) {
    # Returns a Formula string "=<number>" or $null — avoids Windows PS 5.1
    # COM type-marshalling bug when assigning function-returned [double] to Value2.
    if ([string]::IsNullOrWhiteSpace($val) -or $val -eq "NaN") { return $null }
    try {
        $d = [double]$val
        return "=$($d.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture))"
    } catch { return $null }
}

function Set-Value {
    # Assigns a numeric value directly to a cell, avoiding the PS 5.1 COM
    # type-marshalling bug by doing the [double] assignment inside the function body
    # (the bug occurs when a function *returns* [double] for external assignment).
    param($cell, $val)
    if ([string]::IsNullOrWhiteSpace($val) -or $val -eq "NaN") { return }
    try { $cell.Value2 = [double]$val } catch {}
}

function Release-COM($obj) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
}

function Open-Template([string]$name) {
    $path = "$TmplDir\$name"
    if (-not (Test-Path $path)) { throw "Template not found: $path" }
    return $xl.Workbooks.Open($path, 0, $false)
}

# Subscript every "2" in "CO2" occurrences of a chart text object
# (ChartTitle, AxisTitle, DataLabel — anything with .Text and .Characters).
# Only works on static text; formula-linked titles cannot take rich formatting.
function Set-SubscriptCO2($textObj) {
    try {
        $t = [string]$textObj.Text
        $idx = $t.IndexOf("CO2")
        while ($idx -ge 0) {
            $textObj.Characters($idx + 3, 1).Font.Subscript = $true   # Characters is 1-based
            $idx = $t.IndexOf("CO2", $idx + 3)
        }
    } catch {}
}

# Write chart title to Data!A1 and set the chart title as static text
# (static rather than cell-linked so CO2 subscripting can be applied).
function Set-ChartTitle($ws, $chart, [string]$title) {
    $ws.Cells(1,1).Value2 = $title
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $title
    try { $chart.ChartTitle.Font.Name = "Sofia Pro" } catch { try { $chart.ChartTitle.Font.Name = "Calibri" } catch {} }
    try { $chart.ChartTitle.Font.Size = 14 } catch {}
    Set-SubscriptCO2 $chart.ChartTitle
}

# Single-chart figures: record the title in Data!A1 but show no title on the chart.
function Set-DataTitle($ws, $chart, [string]$title) {
    $ws.Cells(1,1).Value2 = $title
    $chart.HasTitle = $false
}

# Copy the first ChartObject from each named chart tab onto one new sheet,
# tiled 2 across (reading order), at each source chart's original size.
# Each panel gets its own short title; an overall title text box sits above,
# and everything is grouped into a single shape so it moves/copies as one canvas.
# Uses Duplicate + Chart.Location rather than clipboard copy/paste — the
# clipboard is unreliable under headless COM (CLIPBRD_E_CANT_OPEN).
function Add-CombinedSheet($wb, [string[]]$chartTabs, [string]$sheetName,
                           [string[]]$panelTitles, [string]$overallTitle) {
    $wsC = $wb.Sheets.Add([System.Type]::Missing, $wb.Sheets($wb.Sheets.Count))
    $wsC.Name = $sheetName
    [double]$pad = 10
    [double]$titleH = 0
    if ($overallTitle) { $titleH = 34 }
    [double]$rowTop = $titleH; [double]$rowH = 0; [double]$firstW = 0
    $shapeNames = @()
    for ($i = 0; $i -lt $chartTabs.Count; $i++) {
        $srcCo = $wb.Sheets($chartTabs[$i]).ChartObjects(1)
        [double]$w = $srcCo.Width; [double]$h = $srcCo.Height
        if ($i -eq 0) { $firstW = $w }
        $dup = $srcCo.Duplicate()
        $null = $dup.Chart.Location(2, $sheetName)   # 2 = xlLocationAsObject
        $co = $wsC.ChartObjects($wsC.ChartObjects().Count)
        $colIdx = $i % 2
        if ($colIdx -eq 0 -and $i -gt 0) { $rowTop += $rowH + $pad; $rowH = 0 }
        $co.Left   = [double]($colIdx * ($w + $pad))
        $co.Top    = $rowTop
        $co.Width  = $w
        $co.Height = $h
        if ($h -gt $rowH) { $rowH = $h }
        # Unique name — duplicated charts otherwise share the source name,
        # which breaks the Shapes.Range grouping below
        try { $co.Name = "Panel_" + ($i + 1) } catch {}
        # Per-panel title (static text so subscripting works)
        if ($panelTitles -and $i -lt $panelTitles.Count) {
            try {
                $chP = $co.Chart
                $chP.HasTitle = $true
                $chP.ChartTitle.Text = $panelTitles[$i]
                try { $chP.ChartTitle.Font.Name = "Sofia Pro" } catch { try { $chP.ChartTitle.Font.Name = "Calibri" } catch {} }
                try { $chP.ChartTitle.Font.Size = 14 } catch {}
                Set-SubscriptCO2 $chP.ChartTitle
            } catch {}
        }
        $shapeNames += $co.Name
    }
    [double]$totW = 2 * $firstW + $pad
    [double]$totH = $rowTop + $rowH
    # Overall title text box spanning both columns
    if ($overallTitle) {
        try {
            $tb = $wsC.Shapes.AddTextbox(1, 0, 0, $totW, $titleH)   # 1 = msoTextOrientationHorizontal
            $tr = $tb.TextFrame2.TextRange
            $tr.Text = $overallTitle
            try { $tr.Font.Name = "Sofia Pro" } catch { try { $tr.Font.Name = "Calibri" } catch {} }
            $tr.Font.Size = 18
            $tr.Font.Bold = -1                                       # msoTrue
            try { $tr.ParagraphFormat.Alignment = 2 } catch {}       # msoAlignCenter
            try { $tb.Line.Visible = 0; $tb.Fill.Visible = 0 } catch {}
            $shapeNames += $tb.Name
        } catch {}
    }
    # White backing rectangle behind the panels — shared canvas look
    try {
        $rect = $wsC.Shapes.AddShape(1, 0, 0, $totW, $totH)          # 1 = msoShapeRectangle
        $rect.Name = "CanvasBg"
        $rect.Fill.ForeColor.RGB = 16777215                          # white
        $rect.Line.Visible = 0                                       # msoFalse
        $rect.ZOrder(1) | Out-Null                                   # 1 = msoSendToBack
        $shapeNames = @("CanvasBg") + $shapeNames
    } catch {}
    # Group background + panels (+ title) into one canvas
    try { $wsC.Shapes.Range([string[]]$shapeNames).Group() | Out-Null } catch {}
    return $wsC
}

# Apply consistent number format and font to a range.
function Format-DataRange($ws, [int]$r1, [int]$r2, [int]$c1, [int]$c2, [string]$numFmt) {
    $cL1 = [char](64 + $c1); $cL2 = [char](64 + $c2)
    $rng = $ws.Range("${cL1}${r1}:${cL2}${r2}")
    $rng.NumberFormat = $numFmt
    try { $rng.Font.Name = "Sofia Pro" } catch { $rng.Font.Name = "Calibri" }
    $rng.Font.Size   = 9
    $rng.Font.Italic = $false
}

# Merge title cell, delete unused template rows below data.
# Deleting rows is faster than clearing cell-by-cell formatting and removes
# all template remnants (colours, borders, values) in one operation.
function Format-Sheet($ws, [int]$startRow, [int]$endRow, [int]$lastCol) {
    $cL = [char](64 + $lastCol)
    try { $ws.Range("A1:${cL}1").Merge() | Out-Null } catch {}
    $delStart = $endRow + 1
    try { $ws.Rows("${delStart}:200").Delete() | Out-Null } catch {}
}

# Import a CSV file into a new sheet of $wb via QueryTable (avoids the
# PowerShell 5.1 bulk-array marshalling bug). Falls back to per-cell writes.
function Import-CsvToTab($wb, [string]$tabName, [string]$csvPath) {
    $wsN = $wb.Sheets.Add($wb.Sheets("Data"))
    $wsN.Name = $tabName
    try {
        $qt = $wsN.QueryTables.Add("TEXT;" + $csvPath, $wsN.Range("A1"))
        $qt.TextFileCommaDelimiter     = $true
        $qt.TextFileDecimalSeparator   = "."
        $qt.TextFileThousandsSeparator = ","
        $qt.PreserveFormatting         = $false
        $qt.RefreshOnFileOpen          = $false
        $qt.Refresh($false) | Out-Null
        $qt.Delete()
        $nRows = $wsN.UsedRange.Rows.Count - 1
        Write-Host "    Tab '${tabName}': $nRows rows (via QueryTable)"
    } catch {
        Write-Host "    WARN: QueryTable failed for '${tabName}': $_"
        $rows = Import-Csv $csvPath
        if ($rows.Count -eq 0) { return }
        $hdrs = @($rows[0].PSObject.Properties.Name)
        for ($c = 0; $c -lt $hdrs.Count; $c++) { $wsN.Cells(1,$c+1).Value2 = $hdrs[$c] }
        $ri = 2
        foreach ($row in $rows) {
            for ($c = 0; $c -lt $hdrs.Count; $c++) {
                $v = $row.($hdrs[$c])
                if (-not [string]::IsNullOrWhiteSpace($v) -and $v -ne "NaN") {
                    try { $wsN.Cells($ri,$c+1).Value2 = [double]$v } catch { $wsN.Cells($ri,$c+1).Value2 = $v }
                }
            }
            $ri++
        }
    }
}

function Save-Close($wb, [string]$outPath) {
    try {
        if (Test-Path $outPath) { Remove-Item $outPath -Force -ErrorAction SilentlyContinue }
        $wb.SaveAs($outPath, 51)   # xlOpenXMLWorkbook
        $wb.Close($false)
        Write-Host "  Saved: $(Split-Path $outPath -Leaf)"
    } catch {
        Write-Host "  WARN: SaveAs failed for $(Split-Path $outPath -Leaf): $_"
        try { $wb.Close($false) } catch {}
    }
}

# Duplicate a Data+Chart sheet pair within the same workbook for a new measure.
# Updates chart series SERIES() formula strings to reference the new data sheet.
function Add-MeasureTab($wb, [string]$srcDataName, [string]$srcChartName,
                        [string]$newDataName, [string]$newChartName) {
    $srcPat  = [regex]::Escape($srcDataName)

    # Copy both sheets to position 1 (Before the first sheet) — this avoids the
    # Windows PS 5.1 COM issue with passing a missing Before arg alongside an After arg.
    # Sheet order within the file is cosmetic; SERIES formulas reference sheets by name.

    $wb.Sheets($srcDataName).Copy($wb.Sheets.Item(1))
    $newDataSheet = $wb.Sheets.Item(1)
    $newDataSheet.Name = $newDataName

    $wb.Sheets($srcChartName).Copy($wb.Sheets.Item(1))
    $newChartSheet = $wb.Sheets.Item(1)
    $newChartSheet.Name = $newChartName

    foreach ($co in $newChartSheet.ChartObjects()) {
        foreach ($s in $co.Chart.SeriesCollection()) {
            $f = $s.Formula
            $f = $f -replace "'?${srcPat}'?!", "${newDataName}!"
            try { $s.Formula = $f } catch {}
        }
    }

    return @($newDataSheet, $newChartSheet)
}
