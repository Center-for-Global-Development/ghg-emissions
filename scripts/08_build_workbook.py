"""
08_build_workbook.py  (v4)
Build Summary_tool.xlsx from pipeline CSVs.

Row layout (Summary): Title | Measure | Start year | End year | Display toggle |
                       spacer | Source headers (frozen) | Data rows

Display toggle (B5): "Absolute (MtCO2e)" / "% of World"
  - Absolute: SUMIFS total for the group/country
  - % of World: group/country divided by World total for the same source
  - Number format switches via CF rule (0.0% when % of World selected)

Sheets:  Summary | Results | Groups (hidden) | Countries_annual (hidden) | Coverage (hidden)
"""

import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.utils import get_column_letter

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
QA_DIR      = os.path.dirname(SCRIPTS_DIR)
OUT_DIR     = os.path.join(QA_DIR, "data", "outputs")

GROUPS_CSV    = os.path.join(OUT_DIR, "groups_summary.csv")
COUNTRIES_CSV = os.path.join(OUT_DIR, "countries_annual.csv")
OUT_XLSX      = os.path.join(OUT_DIR, "Summary_tool_v2.xlsx")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
MEASURE_LABELS = [
    ("co2_excLUC", "CO2 excl. LULUCF"),
    ("co2_incLUC", "CO2 incl. LULUCF"),
    ("ghg_excLUC", "All GHGs excl. LULUCF"),
    ("ghg_incLUC", "All GHGs incl. LULUCF"),
]

SOURCE_ORDER = [
    ("ClimateWatch", "Climate Watch"),
    ("OWID",         "OWID (GCB 2025)"),
    ("PRIMAP",       "PRIMAP-hist v2.7"),
    ("EDGAR",        "EDGAR 2025"),
    ("GCP_fossil",   "GCP fossil"),
    ("GCP_BLUE",     "GCP BLUE"),
    ("GCP_OSCAR",    "GCP OSCAR"),
    ("GCP_LUCE",     "GCP LUCE"),
]

# Sections: (section_label_or_None, [(grp_code, display_name), ...])
GROUPS_SECTIONS = [
    (None, [
        ("World", "World (all countries)"),
    ]),
    ("By institutional group", [
        ("annex_1",        "Annex I"),
        ("annex_2",        "Annex II"),
        ("annex_2_eu",     "Annex II + other EU Members"),
        ("non_annex_2",    "Non-Annex II"),
        ("non_annex_2_eu", "Non-Annex II (excl. EU)"),
        ("oecd",           "OECD"),
        ("eu",             "EU"),
        ("brics",          "BRICS"),
        ("g7",             "G7"),
        ("g20",            "G20"),
    ]),
    ("By income group", [
        ("lic",  "LIC"),
        ("lmic", "LMIC"),
        ("umic", "UMIC"),
        ("hic",  "HIC"),
    ]),
    ("By development group", [
        ("ldc",               "LDC"),
        ("lldc",              "LLDC"),
        ("sids",              "SIDS"),
        ("lic_lmic_sids_ldc", "LIC / LMIC / SIDS / LDC"),
    ]),
    ("By region", [
        ("wb_eca",  "ECA"),
        ("wb_n_am", "North America"),
        ("wb_lac",  "Latin America & Caribbean"),
        ("wb_eap",  "East Asia & Pacific"),
        ("wb_sa",   "South Asia"),
        ("wb_mena", "Middle East & North Africa"),
        ("wb_ssa",  "Sub-Saharan Africa"),
    ]),
]

KEY_COUNTRIES = [
    ("United States",  "USA"),
    ("United Kingdom", "GBR"),
    ("France",         "FRA"),
    ("Germany",        "DEU"),
    ("Italy",          "ITA"),
    ("Japan",          "JPN"),
    ("Canada",         "CAN"),
    ("China",          "CHN"),
    ("India",          "IND"),
    ("Brazil",         "BRA"),
    ("Russia",         "RUS"),
    ("South Africa",   "ZAF"),
]

# Results sheet sections: (section_label, [(grp_code, display_name), ...], note_text_or_None)
# grp_code=None → computed as 1 - annex_1/World
RESULTS_SECTIONS = [
    ("UNFCCC groups", [
        ("annex_2", "Annex II"),
        ("annex_1", "Annex I"),
        (None,      "Non-Annex I"),
    ], None),
    ("Other institutional groups", [
        ("oecd",  "OECD"),
        ("eu",    "EU"),
        ("brics", "BRICS"),
        ("g7",    "G7"),
        ("g20",   "G20"),
    ], None),
    ("Income group", [
        ("hic",  "HIC"),
        ("umic", "UMIC"),
        ("lmic", "LMIC"),
        ("lic",  "LIC"),
    ], (
        "Note: Venezuela (VEN) assigned to UMIC and Ethiopia (ETH) assigned to LIC. "
        "Both are currently WB-unclassified due to data gaps but match their historical income level."
    )),
    ("Development group", [
        ("ldc",               "LDC"),
        ("lldc",              "LLDC"),
        ("sids",              "SIDS"),
        ("lic_lmic_sids_ldc", "LIC / LMIC / SIDS / LDC"),
    ], None),
    ("Regions (World Bank)", [
        ("wb_eca",  "Europe and Central Asia"),
        ("wb_n_am", "North America"),
        ("wb_lac",  "Latin America and the Caribbean"),
        ("wb_mena", "Middle East and North Africa"),
        ("wb_ssa",  "Sub-Saharan Africa"),
        ("wb_eap",  "East Asia and the Pacific"),
        ("wb_sa",   "South Asia"),
    ], (
        "Note: A small number of territories (e.g. Réunion, Guadeloupe, Netherlands Antilles, "
        "Western Sahara) are not assigned to any WB region and are excluded from regional totals "
        "but included in World. Residual is <0.02% of world emissions."
    )),
]

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
CGD_TEAL   = "1A6674"
CGD_AMBER  = "E07B39"
GREY_LIGHT = "F2F2F2"
GREY_MID   = "D9D9D9"
WHITE      = "FFFFFF"
TEXT_GREY  = "808080"
TEXT_MUTED = "777777"
AMBER_PALE = "FFF4EE"

def _f(bold=False, italic=False, color="000000", size=11):
    return Font(bold=bold, italic=italic, color=color, name="Calibri", size=size)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

ALIGN_C = Alignment(horizontal="center", vertical="center")
ALIGN_R = Alignment(horizontal="right",  vertical="center")
ALIGN_L = Alignment(horizontal="left",   vertical="center")

# ---------------------------------------------------------------------------
# Param cell addresses on the Summary sheet
# Row 2: B2 = measure dropdown
# Row 3: B3 = start year
# Row 4: B4 = end year
# Row 5: B5 = display toggle ("Absolute (MtCO2e)" / "% of World")
# Hidden cols P/Q: P2 = measure code (VLOOKUP result); P3:Q6 = lookup table
# ---------------------------------------------------------------------------
MSR_CELL    = "$P$2"
YR_S_CELL   = "$B$3"
YR_E_CELL   = "$B$4"
TOGGLE_CELL = "$B$5"
TOGGLE_PCT  = "% of World"

# Groups tab columns: A=group, B=source, C=measure, D=year, E=value_Mt
GRP_VAL = "Groups!$E:$E"
GRP_GRP = "Groups!$A:$A"
GRP_SRC = "Groups!$B:$B"
GRP_MSR = "Groups!$C:$C"
GRP_YR  = "Groups!$D:$D"

# Countries_annual tab columns: A=iso_code, B=country_name, C=source, D=measure, E=year, F=value_Mt
CA_VAL  = "Countries_annual!$F:$F"
CA_ISO  = "Countries_annual!$A:$A"
CA_SRC  = "Countries_annual!$C:$C"
CA_MSR  = "Countries_annual!$D:$D"
CA_YR   = "Countries_annual!$E:$E"

# Coverage sheet columns: A=source, B=measure, C=min_year, D=max_year (rows 2-50)
COV_SRC  = "Coverage!$A$2:$A$50"
COV_MSR  = "Coverage!$B$2:$B$50"
COV_MINY = "Coverage!$C$2:$C$50"
COV_MAXY = "Coverage!$D$2:$D$50"


def _bare_grp(grp_code, src_code, sheet=""):
    """Raw SUMIFS expression for a group row (no IFERROR wrapper)."""
    m = f"{sheet}{MSR_CELL}"
    s = f"{sheet}{YR_S_CELL}"
    e = f"{sheet}{YR_E_CELL}"
    q = chr(34)
    return (
        f"SUMIFS({GRP_VAL},{GRP_GRP},{q}{grp_code}{q},"
        f"{GRP_SRC},{q}{src_code}{q},"
        f"{GRP_MSR},{m},{GRP_YR},{q}>={q}&{s},"
        f"{GRP_YR},{q}<={q}&{e})"
    )


def _sumifs_grp(grp_code, src_code):
    """Toggle-aware formula for group rows on the Summary sheet.
    Returns absolute total or group/World fraction depending on B5."""
    q = chr(34)
    raw   = _bare_grp(grp_code, src_code)
    world = _bare_grp("World",   src_code)
    return (
        f"=IF({TOGGLE_CELL}={q}{TOGGLE_PCT}{q},"
        f"IFERROR(({raw})/({world}),{q}—{q}),"
        f"IFERROR({raw},{q}—{q}))"
    )


def _sumifs_country(iso_code, src_code):
    """Toggle-aware formula for key-country rows on the Summary sheet."""
    q = chr(34)
    raw = (
        f"SUMIFS({CA_VAL},{CA_ISO},{q}{iso_code}{q},"
        f"{CA_SRC},{q}{src_code}{q},"
        f"{CA_MSR},{MSR_CELL},{CA_YR},{q}>={q}&{YR_S_CELL},"
        f"{CA_YR},{q}<={q}&{YR_E_CELL})"
    )
    world = _bare_grp("World", src_code)
    return (
        f"=IF({TOGGLE_CELL}={q}{TOGGLE_PCT}{q},"
        f"IFERROR(({raw})/({world}),{q}—{q}),"
        f"IFERROR({raw},{q}—{q}))"
    )


def _share_grp(grp_code, src_code):
    """Group / World total for Results sheet (references Summary params)."""
    m = f"Summary!{MSR_CELL}"
    s = f"Summary!{YR_S_CELL}"
    e = f"Summary!{YR_E_CELL}"
    q = chr(34)
    world = (f"SUMIFS({GRP_VAL},{GRP_GRP},{q}World{q},"
             f"{GRP_SRC},{q}{src_code}{q},{GRP_MSR},{m},"
             f"{GRP_YR},{q}>={q}&{s},{GRP_YR},{q}<={q}&{e})")
    grp   = (f"SUMIFS({GRP_VAL},{GRP_GRP},{q}{grp_code}{q},"
             f"{GRP_SRC},{q}{src_code}{q},{GRP_MSR},{m},"
             f"{GRP_YR},{q}>={q}&{s},{GRP_YR},{q}<={q}&{e})")
    return f"=IFERROR(({grp})/({world}),{q}—{q})"


def _share_non_annex1(src_code):
    """(World - Annex I) / World = 1 - annex_1 share."""
    m = f"Summary!{MSR_CELL}"
    s = f"Summary!{YR_S_CELL}"
    e = f"Summary!{YR_E_CELL}"
    q = chr(34)
    world  = (f"SUMIFS({GRP_VAL},{GRP_GRP},{q}World{q},"
              f"{GRP_SRC},{q}{src_code}{q},{GRP_MSR},{m},"
              f"{GRP_YR},{q}>={q}&{s},{GRP_YR},{q}<={q}&{e})")
    annex1 = (f"SUMIFS({GRP_VAL},{GRP_GRP},{q}annex_1{q},"
              f"{GRP_SRC},{q}{src_code}{q},{GRP_MSR},{m},"
              f"{GRP_YR},{q}>={q}&{s},{GRP_YR},{q}<={q}&{e})")
    return f"=IFERROR(1-({annex1})/({world}),{q}—{q})"


def _build_coverage_dict(coverage_df):
    """Build {(source, measure): (min_year, max_year)} for CF formula generation."""
    d = {}
    for row in coverage_df.itertuples(index=False):
        if row.measure in {"co2_excLUC", "co2_incLUC", "ghg_excLUC", "ghg_incLUC"}:
            d[(row.source, row.measure)] = (int(row.min_year), int(row.max_year))
    return d


def _cf_unavailable(src_code, cov_dict):
    """CF formula: TRUE when source has no data for the selected measure.
    Hard-coded measure codes avoid fragile cross-sheet COUNTIFS in CF engine."""
    q = chr(34)
    covered = [m for (s, m) in cov_dict if s == src_code]
    all_m   = ["co2_excLUC", "co2_incLUC", "ghg_excLUC", "ghg_incLUC"]
    missing = [m for m in all_m if m not in covered]
    if not missing:  return "=FALSE"
    if not covered:  return "=TRUE"
    return "=OR(" + ",".join(f'{MSR_CELL}={q}{m}{q}' for m in missing) + ")"


def _cf_partial(src_code, cov_dict):
    """CF formula: TRUE when source covers the measure but not the full year range.
    Hard-coded year bounds avoid fragile cross-sheet MINIFS/MAXIFS in CF engine."""
    q = chr(34)
    parts = []
    for (s, m), (lo, hi) in cov_dict.items():
        if s == src_code:
            parts.append(
                f"AND({MSR_CELL}={q}{m}{q},OR({YR_S_CELL}<{lo},{YR_E_CELL}>{hi}))"
            )
    if not parts:   return "=FALSE"
    return "=OR(" + ",".join(parts) + ")"


def _dxf_font_partial():
    """Minimal differential font for partial coverage (italic + muted colour only)."""
    return Font(italic=True, color=TEXT_MUTED)


def _dxf_font_unavail():
    """Minimal differential font for unavailable (grey colour only, matches fill)."""
    return Font(color=GREY_MID)


def _dxf_fill_unavail():
    """Differential fill for unavailable.
    Excel CF uses bgColor for solid fills in differential styles, not fgColor.
    Setting both ensures correct rendering across Excel versions.
    """
    return PatternFill(patternType="solid", fgColor=GREY_MID, bgColor=GREY_MID)


def _rows_to_cf_range(col_letter, rows):
    """Convert a list of row numbers to a space-separated CF range string."""
    if not rows:
        return f"{col_letter}1"
    ranges, start, prev = [], rows[0], rows[0]
    for r in rows[1:]:
        if r == prev + 1:
            prev = r
        else:
            ranges.append(f"{col_letter}{start}" if start == prev
                          else f"{col_letter}{start}:{col_letter}{prev}")
            start = prev = r
    ranges.append(f"{col_letter}{start}" if start == prev
                  else f"{col_letter}{start}:{col_letter}{prev}")
    return " ".join(ranges)


# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------
print("Loading data ...")
groups           = pd.read_csv(GROUPS_CSV)
countries_annual = pd.read_csv(COUNTRIES_CSV)

actual_sources = set(groups["source"].unique())
actual_groups  = set(groups["group"].unique())
min_year = int(groups["year"].min())
max_year = int(groups["year"].max())

sources_present = [(c, l) for c, l in SOURCE_ORDER if c in actual_sources]
n_src = len(sources_present)

sections_filtered = [
    (name, [(c, l) for c, l in grps if c in actual_groups])
    for name, grps in GROUPS_SECTIONS
]
sections_filtered = [(n, g) for n, g in sections_filtered if g]

# Coverage: min/max year per source×measure (from World group rows)
coverage = (groups[groups["group"] == "World"]
            .groupby(["source", "measure"])
            .agg(min_year=("year", "min"), max_year=("year", "max"))
            .reset_index())

cov_dict = _build_coverage_dict(coverage)
print(f"  Groups: {len(groups):,} rows | {n_src} sources | {len(coverage)} coverage entries")
print(f"  Countries_annual: {len(countries_annual):,} rows")
print(f"  Year range: {min_year}–{max_year}")

# Column layout: A=groups (col 1), B..n_src+1=sources, then key at fixed E/F, hidden lookup at P/Q
KEY_SWATCH_COL  = 5    # col E: key swatch examples
KEY_LABEL_COL   = 6    # col F: key descriptions
LOOKUP_LBL_COL  = 16   # col P: measure labels + formula result (hidden)
LOOKUP_CODE_COL = 17   # col Q: measure codes (hidden)

# ---------------------------------------------------------------------------
# Build workbook
# ---------------------------------------------------------------------------
print("Building workbook ...")
wb = openpyxl.Workbook()
wb.remove(wb.active)

# =========================================================================
# SUMMARY SHEET
# =========================================================================
ws = wb.create_sheet("Summary")
ws.sheet_view.showGridLines = False

# --- Column widths ---
ws.column_dimensions["A"].width = 32
for j in range(2, n_src + 2):
    ws.column_dimensions[get_column_letter(j)].width = 15
ws.column_dimensions["E"].width = 14   # key swatch (col E — overrides source col width)
ws.column_dimensions["F"].width = 36   # key label  (col F — overrides source col width)
ws.column_dimensions[get_column_letter(LOOKUP_LBL_COL)].hidden  = True
ws.column_dimensions[get_column_letter(LOOKUP_CODE_COL)].hidden = True

# --- Row 1: Title ---
ws["A1"] = "GHG Emissions Comparison"
ws["A1"].font = _f(bold=True, size=14, color=CGD_TEAL)

# --- Row 2: Measure parameter ---
ws["A2"] = "Measure:"
ws["A2"].font = _f(bold=True)
ws["B2"] = "CO2 excl. LULUCF"
ws["B2"].font = _f()
ws["B2"].fill = _fill(GREY_LIGHT)

# --- Row 3: Start year ---
ws["A3"] = "Start year:"
ws["A3"].font = _f(bold=True)
ws["B3"] = 1990
ws["B3"].font = _f()
ws["B3"].fill = _fill(GREY_LIGHT)
ws["B3"].number_format = "0"

# --- Row 4: End year ---
ws["A4"] = "End year:"
ws["A4"].font = _f(bold=True)
ws["B4"] = max_year
ws["B4"].font = _f()
ws["B4"].fill = _fill(GREY_LIGHT)
ws["B4"].number_format = "0"

# --- Key legend (rows 2-4, to the right of params) ---
key_hdr       = ws.cell(row=2, column=KEY_SWATCH_COL)
key_hdr.value = "Key:"
key_hdr.font  = _f(bold=True, color=TEXT_GREY)

# Partial coverage: italic, muted text
k3       = ws.cell(row=3, column=KEY_SWATCH_COL)
k3.value = "1,234.5"
k3.font  = _f(italic=True, color=TEXT_MUTED)
k3.alignment = ALIGN_R
k3_lbl       = ws.cell(row=3, column=KEY_LABEL_COL)
k3_lbl.value = "Partial coverage — year range not fully covered by this source"
k3_lbl.font  = _f(color=TEXT_GREY)

# Unavailable: grey fill, invisible text
k4       = ws.cell(row=4, column=KEY_SWATCH_COL)
k4.value = " "
k4.fill  = _fill(GREY_MID)
k4.font  = _f(color=GREY_MID)
k4_lbl       = ws.cell(row=4, column=KEY_LABEL_COL)
k4_lbl.value = "Unavailable — measure or year range not covered by this source"
k4_lbl.font  = _f(color=TEXT_GREY)

# --- Hidden measure lookup: labels in col P, codes in col Q; formula result in P2 ---
p = get_column_letter(LOOKUP_LBL_COL)
q = get_column_letter(LOOKUP_CODE_COL)
for i, (code, label) in enumerate(MEASURE_LABELS, start=3):
    ws.cell(row=i, column=LOOKUP_LBL_COL).value  = label
    ws.cell(row=i, column=LOOKUP_CODE_COL).value = code
ws.cell(row=2, column=LOOKUP_LBL_COL).value = f"=VLOOKUP($B$2,${p}$3:${q}$6,2,0)"

# --- Data validations ---
measure_list = ",".join(lbl for _, lbl in MEASURE_LABELS)
dv_m = DataValidation(type="list", formula1=f'"{measure_list}"', allow_blank=False)
dv_m.sqref = "B2"
ws.add_data_validation(dv_m)

dv_s = DataValidation(
    type="whole", operator="between",
    formula1=str(min_year), formula2="$B$4",
    showErrorMessage=True,
    errorTitle="Invalid start year",
    error=f"Start year must be between {min_year} and End year."
)
dv_s.sqref = "B3"
ws.add_data_validation(dv_s)

dv_e = DataValidation(
    type="whole", operator="between",
    formula1="$B$3", formula2=str(max_year),
    showErrorMessage=True,
    errorTitle="Invalid end year",
    error=f"End year must be between Start year and {max_year}."
)
dv_e.sqref = "B4"
ws.add_data_validation(dv_e)

# --- Row 5: Display toggle ---
ws["A5"] = "Display:"
ws["A5"].font = _f(bold=True)
ws["B5"] = "Absolute (MtCO2e)"
ws["B5"].font = _f()
ws["B5"].fill = _fill(GREY_LIGHT)

display_list = f'"Absolute (MtCO2e),{TOGGLE_PCT}"'
dv_d = DataValidation(type="list", formula1=display_list, allow_blank=False)
dv_d.sqref = "B5"
ws.add_data_validation(dv_d)

# --- Row 6: thin spacer ---
ws.row_dimensions[6].height = 5

# --- Row 7: Source column headers ---
teal_fill = _fill(CGD_TEAL)
ws.cell(row=7, column=1).value     = "Group"
ws.cell(row=7, column=1).font      = _f(bold=True, color=WHITE)
ws.cell(row=7, column=1).fill      = teal_fill
ws.cell(row=7, column=1).alignment = ALIGN_L
for j, (_, src_label) in enumerate(sources_present, start=2):
    c = ws.cell(row=7, column=j)
    c.value     = src_label
    c.font      = _f(bold=True, color=WHITE)
    c.fill      = teal_fill
    c.alignment = ALIGN_C

# --- Freeze rows 1-7 ---
ws.freeze_panes = "A8"

# --- Data rows ---
current_row = 8
data_rows   = []

for section_name, grp_list in sections_filtered:
    if section_name is not None:
        c = ws.cell(row=current_row, column=1)
        c.value = section_name
        c.font  = _f(bold=True, color=CGD_AMBER)
        for j in range(1, n_src + 2):
            ws.cell(row=current_row, column=j).fill = _fill(AMBER_PALE)
        current_row += 1

    for grp_code, grp_label in grp_list:
        is_world = (grp_code == "World")
        c = ws.cell(row=current_row, column=1)
        c.value     = grp_label
        c.font      = _f(bold=is_world)
        c.alignment = ALIGN_L
        if is_world:
            for j in range(1, n_src + 2):
                ws.cell(row=current_row, column=j).fill = _fill(GREY_LIGHT)

        for j, (src_code, _) in enumerate(sources_present, start=2):
            cell = ws.cell(row=current_row, column=j)
            cell.value         = _sumifs_grp(grp_code, src_code)
            cell.number_format = "#,##0"
            cell.alignment     = ALIGN_R

        data_rows.append(current_row)
        current_row += 1

    current_row += 1  # blank spacer between sections

# Key countries section
c = ws.cell(row=current_row, column=1)
c.value = "Key countries"
c.font  = _f(bold=True, color=CGD_AMBER)
for j in range(1, n_src + 2):
    ws.cell(row=current_row, column=j).fill = _fill(AMBER_PALE)
current_row += 1

for country_name, iso in KEY_COUNTRIES:
    ws.cell(row=current_row, column=1).value     = country_name
    ws.cell(row=current_row, column=1).font      = _f()
    ws.cell(row=current_row, column=1).alignment = ALIGN_L
    for j, (src_code, _) in enumerate(sources_present, start=2):
        cell = ws.cell(row=current_row, column=j)
        cell.value         = _sumifs_country(iso, src_code)
        cell.number_format = "#,##0"
        cell.alignment     = ALIGN_R
    data_rows.append(current_row)
    current_row += 1

# --- Conditional formatting (per source column, data rows only) ---
# Rules are added in priority order: first added = priority 1 = fires first.
# Order: (1) % format [fires for all data when toggle set],
#         (2) unavailable [grey fill, stopIfTrue],
#         (3) partial [italic muted].
pct_dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=164, formatCode="0.0%"))

for j, (src_code, _) in enumerate(sources_present, start=2):
    col_ltr  = get_column_letter(j)
    cf_range = _rows_to_cf_range(col_ltr, data_rows)

    # Rule 1: switch to % format when display toggle = "% of World"
    ws.conditional_formatting.add(
        cf_range,
        Rule(type="expression", dxf=pct_dxf,
             formula=[f'={TOGGLE_CELL}="{TOGGLE_PCT}"'])
    )
    # Rule 2: completely unavailable → grey fill + invisible text (stopIfTrue)
    ws.conditional_formatting.add(
        cf_range,
        FormulaRule(formula=[_cf_unavailable(src_code, cov_dict)],
                    fill=_dxf_fill_unavail(),
                    font=_dxf_font_unavail(),
                    stopIfTrue=True)
    )
    # Rule 3: partial coverage → italic, muted text
    ws.conditional_formatting.add(
        cf_range,
        FormulaRule(formula=[_cf_partial(src_code, cov_dict)],
                    font=_dxf_font_partial())
    )

# =========================================================================
# RESULTS SHEET
# =========================================================================
ws_r = wb.create_sheet("Results")
ws_r.sheet_view.showGridLines = False
ws_r.column_dimensions["A"].width = 38
for j in range(2, n_src + 2):
    ws_r.column_dimensions[get_column_letter(j)].width = 15

ws_r["A1"] = "Share of World total"
ws_r["A1"].font = _f(bold=True, size=14, color=CGD_TEAL)

ws_r["A2"] = "Measure:"
ws_r["A2"].font = _f(bold=True)
ws_r["B2"] = "=Summary!B2"
ws_r["B2"].font = _f()

ws_r["D2"] = "Years:"
ws_r["D2"].font = _f(bold=True)
ws_r["E2"] = "=Summary!B3&\" \u2013 \"&Summary!B4"
ws_r["E2"].font = _f()

ws_r.row_dimensions[3].height = 5

ws_r.cell(row=4, column=1).value     = "Group"
ws_r.cell(row=4, column=1).font      = _f(bold=True, color=WHITE)
ws_r.cell(row=4, column=1).fill      = _fill(CGD_TEAL)
ws_r.cell(row=4, column=1).alignment = ALIGN_L
for j, (_, src_label) in enumerate(sources_present, start=2):
    c = ws_r.cell(row=4, column=j)
    c.value     = src_label
    c.font      = _f(bold=True, color=WHITE)
    c.fill      = _fill(CGD_TEAL)
    c.alignment = ALIGN_C

ws_r.freeze_panes = "A5"

res_row = 5
for section_label, grp_list, note_text in RESULTS_SECTIONS:
    # Section header row (amber)
    for j in range(1, n_src + 2):
        ws_r.cell(row=res_row, column=j).fill = _fill(AMBER_PALE)
    c = ws_r.cell(row=res_row, column=1)
    c.value     = section_label
    c.font      = _f(bold=True, color=CGD_AMBER)
    c.alignment = ALIGN_L
    res_row += 1

    # Data rows
    for grp_code, grp_label in grp_list:
        ws_r.cell(row=res_row, column=1).value     = grp_label
        ws_r.cell(row=res_row, column=1).font      = _f()
        ws_r.cell(row=res_row, column=1).alignment = ALIGN_L
        for j, (src_code, _) in enumerate(sources_present, start=2):
            c = ws_r.cell(row=res_row, column=j)
            c.value         = (_share_non_annex1(src_code)
                               if grp_code is None
                               else _share_grp(grp_code, src_code))
            c.number_format = "0.0%"
            c.alignment     = ALIGN_R
        res_row += 1

    # Optional note row
    if note_text:
        c = ws_r.cell(row=res_row, column=1)
        c.value     = note_text
        c.font      = _f(italic=True, color=TEXT_MUTED, size=9)
        c.alignment = ALIGN_L
        ws_r.row_dimensions[res_row].height = 28
        res_row += 1

    res_row += 1  # blank spacer between sections

# =========================================================================
# COVERAGE SHEET (hidden — used by CF formulas)
# =========================================================================
ws_cov = wb.create_sheet("Coverage")
for j, h in enumerate(["source", "measure", "min_year", "max_year"], start=1):
    ws_cov.cell(row=1, column=j).value = h
    ws_cov.cell(row=1, column=j).font  = _f(bold=True)
for i, row in enumerate(coverage.itertuples(index=False), start=2):
    ws_cov.cell(row=i, column=1).value = row.source
    ws_cov.cell(row=i, column=2).value = row.measure
    ws_cov.cell(row=i, column=3).value = int(row.min_year)
    ws_cov.cell(row=i, column=4).value = int(row.max_year)
ws_cov.sheet_state = "hidden"

# =========================================================================
# GROUPS DATA TAB (hidden)
# A=group, B=source, C=measure, D=year, E=value_Mt
# =========================================================================
print("Writing Groups data tab ...")
ws_grp = wb.create_sheet("Groups")
hdr = ["group", "source", "measure", "year", "value_Mt"]
for j, h in enumerate(hdr, start=1):
    ws_grp.cell(row=1, column=j).value = h
    ws_grp.cell(row=1, column=j).font  = _f(bold=True)
for row in groups[hdr].itertuples(index=False):
    ws_grp.append(list(row))
ws_grp.sheet_state = "hidden"
print(f"  Written {len(groups):,} rows")

# =========================================================================
# COUNTRIES_ANNUAL DATA TAB (hidden)
# A=iso_code, B=country_name, C=source, D=measure, E=year, F=value_Mt
# =========================================================================
print("Writing Countries_annual data tab ...")
ws_ca = wb.create_sheet("Countries_annual")
ca_hdr = ["iso_code", "country_name", "source", "measure", "year", "value_Mt"]
for j, h in enumerate(ca_hdr, start=1):
    ws_ca.cell(row=1, column=j).value = h
    ws_ca.cell(row=1, column=j).font  = _f(bold=True)
for row in countries_annual[ca_hdr].itertuples(index=False):
    ws_ca.append(list(row))
ws_ca.sheet_state = "hidden"
print(f"  Written {len(countries_annual):,} rows")

# =========================================================================
# Save
# =========================================================================
wb.active = wb["Summary"]
print(f"Saving {os.path.basename(OUT_XLSX)} ...")
wb.save(OUT_XLSX)
print(f"Done.  {len(data_rows)} data rows | {len(sources_present)} sources")
