<#
.SYNOPSIS
_chart_multiline.ps1 — MultiLine band-chart helpers (Figs D, E, K, F).

Dot-source from any chart script that uses the MultiLine template:
    . "$PSScriptRoot\_chart_multiline.ps1"

Provides four functions:
  Build-MultiLineVals  — populate vals_[label] tab (raw annual values + cumulative calcs)
  Build-MultiLineCalcs — populate data_[label] tab (chart-facing cols only, references vals)
  Build-FigKData       — Fig K variant: annual % already in CSV, no vals tab needed
  Add-GMSTLines        — add Annex II / Annex I GMST solid-line overlay to a chart

Requires Set-Value from _chart_helpers.ps1.

Tab layout (data_[label]):
  A = year, B-H = band aggregation (from I-P), I-L = pct_annex2 per source,
  M-P = pct_annex1 per source, Q = annex2_GMST, R = annex1_GMST.
  Five-series stacked area: _Base(B), Pure_AII(F), Overlap(G), AI_Gap(H), Pure_AI(S).
  Stack total = AI_Max (col E) regardless of band crossing.
#>

$cgdGMST_AII = 255        # red (BGR)
$cgdGMST_AI  = 8388608    # dark blue (BGR)

# Populate vals_[label] tab with annual MtCO2e values AND calc columns.
#
# Vals tab layout:
#   A=Year
#   B-M: raw annual values, 3 cols per source slot × 4 slots (annex2, annex1, non_annex1)
#   N+:  calc cols, 6 cols per source slot × 4 slots:
#          cum_annex2, cum_annex1, cum_non_ann1, cum_world (=ann1+non_ann1),
#          pct_annex2 (=cum_annex2/cum_world*100), pct_annex1 (=cum_annex1/cum_world*100)
#
# All cumulative formulas expand in the appropriate direction (forward for FigD,
# reverse for FigE) using R1C1 notation.
function Build-MultiLineVals($wsV, $valsName, $figKey, $csvRows, $cm) {
    try { $null = $wsV.Cells.UnMerge() } catch {}
    $null = $wsV.Cells.ClearContents()

    # End row = row 3 (first data row) + number of data rows - 1
    $endRowInt = 2 + @($csvRows).Count
    if ($endRowInt -lt 3) { return $endRowInt }

    # Row 2: headers for raw value cols
    $wsV.Cells(2, 1).Value2 = "Year"
    for ($i = 0; $i -lt 4; $i++) {
        $a2col = $cm.A[$i]
        if ($a2col -eq "") { continue }
        $src   = $a2col -replace "^annex2_", ""
        $cBase = 2 + 3 * $i
        $wsV.Cells(2, $cBase  ).Value2 = "annex2_${src}"
        $wsV.Cells(2, $cBase+1).Value2 = "annex1_${src}"
        $wsV.Cells(2, $cBase+2).Value2 = "non_annex1_${src}"
    }
    # Row 2: headers for calc cols (starting col 14)
    for ($i = 0; $i -lt 4; $i++) {
        $a2col = $cm.A[$i]
        if ($a2col -eq "") { continue }
        $src   = $a2col -replace "^annex2_", ""
        $cCalc = 14 + 6 * $i
        $wsV.Cells(2, $cCalc  ).Value2 = "cum_annex2_${src}"
        $wsV.Cells(2, $cCalc+1).Value2 = "cum_annex1_${src}"
        $wsV.Cells(2, $cCalc+2).Value2 = "cum_non_ann1_${src}"
        $wsV.Cells(2, $cCalc+3).Value2 = "cum_world_${src}"
        $wsV.Cells(2, $cCalc+4).Value2 = "pct_annex2_${src}"
        $wsV.Cells(2, $cCalc+5).Value2 = "pct_annex1_${src}"
    }

    # Data rows: write raw values from CSV
    $r = 3
    foreach ($row in $csvRows) {
        $wsV.Cells($r, 1).Value2 = [int]$row.year
        for ($i = 0; $i -lt 4; $i++) {
            $a2col = $cm.A[$i]; $a1col = $cm.B[$i]; $nacol = $cm.C[$i]
            if ($a2col -eq "") { continue }
            $cBase = 2 + 3 * $i
            Set-Value $wsV.Cells($r, $cBase  ) $row.$a2col
            Set-Value $wsV.Cells($r, $cBase+1) $row.$a1col
            Set-Value $wsV.Cells($r, $cBase+2) $row.$nacol
        }
        $r++
    }

    # Calc cols: cumulative sum formulas (FormulaR1C1, one range-fill per column)
    for ($i = 0; $i -lt 4; $i++) {
        $a2col = $cm.A[$i]
        if ($a2col -eq "") { continue }
        $cBase = 2 + 3 * $i   # annex2 raw col in THIS tab
        $vA2   = $cBase        # same sheet, so just col number
        $vA1   = $cBase + 1
        $vNA1  = $cBase + 2
        $cCalc = 14 + 6 * $i

        $rngA2  = $wsV.Range($wsV.Cells(3,$cCalc),   $wsV.Cells($endRowInt,$cCalc))
        $rngA1  = $wsV.Range($wsV.Cells(3,$cCalc+1), $wsV.Cells($endRowInt,$cCalc+1))
        $rngNA1 = $wsV.Range($wsV.Cells(3,$cCalc+2), $wsV.Cells($endRowInt,$cCalc+2))
        $rngW   = $wsV.Range($wsV.Cells(3,$cCalc+3), $wsV.Cells($endRowInt,$cCalc+3))
        $rngP2  = $wsV.Range($wsV.Cells(3,$cCalc+4), $wsV.Cells($endRowInt,$cCalc+4))
        $rngP1  = $wsV.Range($wsV.Cells(3,$cCalc+5), $wsV.Cells($endRowInt,$cCalc+5))

        if ($figKey -eq "fig6") {
            # Forward expanding sum: fixed start (R3), expanding end
            $rngA2.FormulaR1C1  = "=IFERROR(SUM(R3C${vA2}:RC${vA2}),"""")"
            $rngA1.FormulaR1C1  = "=IFERROR(SUM(R3C${vA1}:RC${vA1}),"""")"
            $rngNA1.FormulaR1C1 = "=IFERROR(SUM(R3C${vNA1}:RC${vNA1}),"""")"
        } else {
            # Reverse sum: expanding start, fixed end.
            # Guard with ISBLANK: if the current start-year row has no raw data (source not
            # yet in coverage), return NA() immediately rather than summing the later years —
            # that would produce a constant flat line for all pre-coverage start years.
            $rngA2.FormulaR1C1  = "=IF(ISBLANK(RC${vA2}),NA(),IFERROR(SUM(RC${vA2}:R${endRowInt}C${vA2}),NA()))"
            $rngA1.FormulaR1C1  = "=IF(ISBLANK(RC${vA1}),NA(),IFERROR(SUM(RC${vA1}:R${endRowInt}C${vA1}),NA()))"
            $rngNA1.FormulaR1C1 = "=IF(ISBLANK(RC${vNA1}),NA(),IFERROR(SUM(RC${vNA1}:R${endRowInt}C${vNA1}),NA()))"
        }
        # cum_world = cum_annex1 + cum_non_annex1  (explicit two-step derivation)
        $rngW.FormulaR1C1  = "=IFERROR(RC[-2]+RC[-1],"""")"
        # pct = cumulative group / cum_world * 100.
        # Use NA() fallback so pre-coverage years (cum_world=0 → #DIV/0!) propagate as
        # #N/A to the data tab, where charts render gaps rather than zeros.
        $rngP2.FormulaR1C1 = "=IFERROR(RC[-4]/RC[-1]*100,NA())"
        $rngP1.FormulaR1C1 = "=IFERROR(RC[-4]/RC[-2]*100,NA())"
    }

    return $endRowInt
}

# Write data_[label] tab: chart-facing cols A-R only, referencing vals tab for pct values.
# Left region: A=year, B-H=band aggregation, I-L=pct_annex2 refs, M-P=pct_annex1 refs,
#              Q=annex2_GMST, R=annex1_GMST.
# No calc columns in this tab — all intermediate calcs are in vals_[label].
function Build-MultiLineCalcs($wsD, $valsTabName, $cm, $endRow, $csvRows) {
    try { $null = $wsD.Cells.UnMerge() } catch {}
    $null = $wsD.Range("A2:S200").ClearContents()

    $endRowInt = [int]$endRow
    if ($endRowInt -lt 3) { return }

    # Row 2: headers
    $wsD.Cells(2,  1).Value2 = "Year"
    $wsD.Cells(2,  2).Value2 = "AII_Min"
    $wsD.Cells(2,  3).Value2 = "AII_Max"
    $wsD.Cells(2,  4).Value2 = "AI_Min"
    $wsD.Cells(2,  5).Value2 = "AI_Max"
    $wsD.Cells(2,  6).Value2 = "Pure_AII"
    $wsD.Cells(2,  7).Value2 = "Overlap"
    $wsD.Cells(2,  8).Value2 = "AI_Gap"
    $wsD.Cells(2, 17).Value2 = "Pure_AI"
    $wsD.Cells(2, 18).Value2 = "annex2_GMST"
    $wsD.Cells(2, 19).Value2 = "annex1_GMST"
    for ($i = 0; $i -lt 4; $i++) {
        $a2col = $cm.A[$i]
        if ($a2col -eq "") { continue }
        $src = $a2col -replace "^annex2_", ""
        $wsD.Cells(2,  9 + $i).Value2 = "pct_annex2_${src}"
        $wsD.Cells(2, 13 + $i).Value2 = "pct_annex1_${src}"
    }

    # Col A: year (reference vals tab col A)
    $wsD.Range($wsD.Cells(3,1), $wsD.Cells($endRowInt,1)).FormulaR1C1 = "='${valsTabName}'!RC1"

    # Cols I-P: reference pct cols from vals tab (cols 18+6i+4 and 18+6i+5 = pct_annex2/1)
    # vals tab calc cols start at col 14 (0-indexed slot i): pct_annex2 = col 14+6i+4, pct_annex1 = 14+6i+5
    # Direct reference (no IFERROR here): vals tab pct already returns NA() for pre-coverage
    # years (0/0 = #DIV/0 → NA()), which propagates here and charts render as gaps not zeros.
    for ($i = 0; $i -lt 4; $i++) {
        $a2col = $cm.A[$i]
        if ($a2col -eq "") { continue }
        $vPctA2 = 14 + 6 * $i + 4   # pct_annex2 col in vals tab
        $vPctA1 = $vPctA2 + 1        # pct_annex1 col in vals tab
        $lA2Col = 9  + $i            # I-L in data tab
        $lA1Col = 13 + $i            # M-P in data tab
        $wsD.Range($wsD.Cells(3,$lA2Col), $wsD.Cells($endRowInt,$lA2Col)).FormulaR1C1 = "='${valsTabName}'!RC${vPctA2}"
        $wsD.Range($wsD.Cells(3,$lA1Col), $wsD.Cells($endRowInt,$lA1Col)).FormulaR1C1 = "='${valsTabName}'!RC${vPctA1}"
    }

    # Cols B-H, S: band aggregation (MIN/MAX of I-L=9-12 and M-P=13-16)
    # B = AII_Min, C = AII_Max, D = AI_Min, E = AI_Max
    # 5-series stacked area: _Base(B) + Pure_AII(F) + Overlap(G) + AI_Gap(H) + Pure_AI(S)
    #   Pure_AII = lower gold zone = MIN(C,D) - B  (from AII_Min up to MIN of AII_Max/AI_Min)
    #   Overlap  = olive zone     = MAX(0, C-D)     (zone where both bands are present; 0 if no crossing)
    #   AI_Gap   = transparent    = MAX(0, D-C)     (gap between bands when no crossing; 0 if crossing)
    #   Pure_AI  = upper teal zone = MAX(0, E-MAX(C,D))  (from MAX of AII_Max/AI_Min up to AI_Max)
    # Stack top always = E = AI_Max. Verified: B + F + G + H + S = AI_Max in both crossing/non-crossing.
    # AGGREGATE(5/4, 6, range) = MIN/MAX ignoring errors — handles NA() from partial-coverage sources.
    $wsD.Range($wsD.Cells(3,2), $wsD.Cells($endRowInt,2)).FormulaR1C1 = "=IFERROR(AGGREGATE(5,6,RC[7]:RC[10]),"""")"  # AII_Min
    $wsD.Range($wsD.Cells(3,3), $wsD.Cells($endRowInt,3)).FormulaR1C1 = "=IFERROR(AGGREGATE(4,6,RC[6]:RC[9]),"""")"   # AII_Max
    $wsD.Range($wsD.Cells(3,4), $wsD.Cells($endRowInt,4)).FormulaR1C1 = "=IFERROR(AGGREGATE(5,6,RC[9]:RC[12]),"""")"  # AI_Min
    $wsD.Range($wsD.Cells(3,5), $wsD.Cells($endRowInt,5)).FormulaR1C1 = "=IFERROR(AGGREGATE(4,6,RC[8]:RC[11]),"""")"  # AI_Max
    $wsD.Range($wsD.Cells(3,6), $wsD.Cells($endRowInt,6)).FormulaR1C1 = "=IFERROR(MIN(RC[-3],RC[-2])-RC[-4],"""")"   # Pure_AII = MIN(C,D)-B
    $wsD.Range($wsD.Cells(3,7), $wsD.Cells($endRowInt,7)).FormulaR1C1 = "=IFERROR(MAX(0,RC[-4]-RC[-3]),"""")"         # Overlap = MAX(0,C-D)
    $wsD.Range($wsD.Cells(3,8), $wsD.Cells($endRowInt,8)).FormulaR1C1 = "=IFERROR(MAX(0,RC[-4]-RC[-5]),"""")"         # AI_Gap = MAX(0,D-C)
    # Pure_AI in col Q(17) — matches template series 5 which references col Q.
    # Keeping Pure_AI here is critical: overwriting Q with GMST data (as this code
    # previously did) causes the template's series 5 to plot GMST as a stacked area,
    # pushing the chart stack far above 100%.
    $wsD.Range($wsD.Cells(3,17), $wsD.Cells($endRowInt,17)).FormulaR1C1 = "=IFERROR(MAX(0,RC5-MAX(RC3,RC4)),"""")"    # Pure_AI = MAX(0,E-MAX(C,D))

    # GMST columns R(18) and S(19): plain values from CSV
    $r = 3
    foreach ($row in $csvRows) {
        Set-Value $wsD.Cells($r, 18) $row.annex2_GMST
        Set-Value $wsD.Cells($r, 19) $row.annex1_GMST
        $r++
    }
}

# Populate a data_[label] tab for Fig K with annual % values already in the CSV.
# No vals tab needed — script 09 already computed the percentages.
# Writes pct values to I-L (annex2) and M-P (annex1); band formulas B-H and col S computed here.
function Build-FigKData($wsD, $cm, $csvRows) {
    try { $null = $wsD.Cells.UnMerge() } catch {}
    $null = $wsD.Range("A2:S200").ClearContents()

    $endRowInt = 2 + @($csvRows).Count
    if ($endRowInt -lt 3) { return $endRowInt }

    # Headers
    $wsD.Cells(2,  1).Value2 = "Year"
    $wsD.Cells(2,  2).Value2 = "AII_Min"
    $wsD.Cells(2,  3).Value2 = "AII_Max"
    $wsD.Cells(2,  4).Value2 = "AI_Min"
    $wsD.Cells(2,  5).Value2 = "AI_Max"
    $wsD.Cells(2,  6).Value2 = "Pure_AII"
    $wsD.Cells(2,  7).Value2 = "Overlap"
    $wsD.Cells(2,  8).Value2 = "AI_Gap"
    $wsD.Cells(2, 17).Value2 = "Pure_AI"
    for ($i = 0; $i -lt 4; $i++) {
        $a2col = $cm.A[$i]
        if ($a2col -eq "") { continue }
        $src = $a2col -replace "^annex2_", ""
        $wsD.Cells(2,  9 + $i).Value2 = "pct_annex2_${src}"
        $wsD.Cells(2, 13 + $i).Value2 = "pct_annex1_${src}"
    }

    # Write annual pct values directly from CSV to I-L (annex2) and M-P (annex1)
    $r = 3
    foreach ($row in $csvRows) {
        $wsD.Cells($r, 1).Value2 = [int]$row.year
        for ($i = 0; $i -lt 4; $i++) {
            $a2col = $cm.A[$i]; $a1col = $cm.B[$i]
            if ($a2col -ne "") { Set-Value $wsD.Cells($r, 9+$i) $row.$a2col }
            if ($a1col -ne "") { Set-Value $wsD.Cells($r, 13+$i) $row.$a1col }
        }
        $r++
    }

    # Col A: year reference; B-E: band aggregation over I-L and M-P
    $wsD.Range($wsD.Cells(3,2), $wsD.Cells($endRowInt,2)).FormulaR1C1 = "=IFERROR(MIN(RC[7]:RC[10]),"""")"
    $wsD.Range($wsD.Cells(3,3), $wsD.Cells($endRowInt,3)).FormulaR1C1 = "=IFERROR(MAX(RC[6]:RC[9]),"""")"
    $wsD.Range($wsD.Cells(3,4), $wsD.Cells($endRowInt,4)).FormulaR1C1 = "=IFERROR(MIN(RC[9]:RC[12]),"""")"
    $wsD.Range($wsD.Cells(3,5), $wsD.Cells($endRowInt,5)).FormulaR1C1 = "=IFERROR(MAX(RC[8]:RC[11]),"""")"
    # Band height formulas (same logic as Build-MultiLineCalcs)
    $wsD.Range($wsD.Cells(3,6), $wsD.Cells($endRowInt,6)).FormulaR1C1 = "=IFERROR(MIN(RC[-3],RC[-2])-RC[-4],"""")"
    $wsD.Range($wsD.Cells(3,7), $wsD.Cells($endRowInt,7)).FormulaR1C1 = "=IFERROR(MAX(0,RC[-4]-RC[-3]),"""")"
    $wsD.Range($wsD.Cells(3,8), $wsD.Cells($endRowInt,8)).FormulaR1C1 = "=IFERROR(MAX(0,RC[-4]-RC[-5]),"""")"
    # Pure_AI in col Q(17) — matches template series 5 which references col Q
    $wsD.Range($wsD.Cells(3,17), $wsD.Cells($endRowInt,17)).FormulaR1C1 = "=IFERROR(MAX(0,RC5-MAX(RC3,RC4)),"""")"

    return $endRowInt
}

# Add GMST solid line series to a Chart_dashes chart.
function Add-GMSTLines($chML, $wsML, $startRowML, $endRowML) {
    $dn = $wsML.Name
    # GMST labels in header row (cols R=18, S=19 — col Q=17 is Pure_AI, reserved for template series 5)
    $wsML.Cells(2, 18).Value2 = "Annex II GMST"
    $wsML.Cells(2, 19).Value2 = "Annex I GMST"

    # Annex II GMST: solid amber-red line, weight 2.5 (bolder than source lines)
    $sG1 = $chML.SeriesCollection().NewSeries()
    $sG1.Name      = "='${dn}'!`$R`$2"
    $sG1.ChartType  = 4       # xlLine
    $sG1.AxisGroup  = 1
    $sG1.MarkerStyle = -4142
    $sG1.Values    = $wsML.Range("R${startRowML}:R${endRowML}")
    $sG1.XValues   = $wsML.Range("A${startRowML}:A${endRowML}")
    $sG1.Format.Line.ForeColor.RGB = $cgdGMST_AII
    $sG1.Format.Line.Weight        = 2.5
    try { $sG1.Format.Line.DashStyle = 1 } catch {}   # Solid

    # Annex I GMST: solid dark blue line
    $sG2 = $chML.SeriesCollection().NewSeries()
    $sG2.Name      = "='${dn}'!`$S`$2"
    $sG2.ChartType  = 4
    $sG2.AxisGroup  = 1
    $sG2.MarkerStyle = -4142
    $sG2.Values    = $wsML.Range("S${startRowML}:S${endRowML}")
    $sG2.XValues   = $wsML.Range("A${startRowML}:A${endRowML}")
    $sG2.Format.Line.ForeColor.RGB = $cgdGMST_AI
    $sG2.Format.Line.Weight        = 2.5
    try { $sG2.Format.Line.DashStyle = 1 } catch {}   # Solid
}
