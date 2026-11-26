"""Build a polished, Google-Sheets-ready Cost Per Use workbook (.xlsx).

Generates four tabs with cohesive typography, visual hierarchy, and native charts:
1. 'Quick Calc': Minimal single-item Cost Per Use calculator card.
2. 'Calculator & Compare': Executive KPI cards, 3-option comparison, break-even analysis,
   plus live Cost-per-Use bar chart and Cumulative Replacement Cost crossover line chart.
3. 'Master Tracker': 100-row purchase & durability tracker with KPI strip, status/tier pills,
   zebra striping, color scale, and clean horizontal-only freeze panes.
4. 'Insights & Charts': Live category, ownership status, and value-tier rollups with 4 charts.
"""

from pathlib import Path
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


CATEGORIES = [
    "Apparel",
    "Tech & Gear",
    "Home & Furniture",
    "Tools",
    "Vehicle",
    "Subscription / Service",
    "Hobby",
    "Other",
]
FREQ_UNITS = ["Day", "Week", "Month", "Year"]
STATUSES = ["Considering", "Owned", "Retired", "Skipped"]
RATINGS = ["1", "2", "3", "4", "5"]

# Cohesive Slate / Cobalt / Emerald editorial palette
C_CANVAS = "F8FAFC"          # Slate 50 (outer dashboard surface)
C_WHITE = "FFFFFF"           # Pure white cards & tables
C_ZEBRA = "F8FAFC"           # Subtle alternating row
C_HEADER_DARK = "0F172A"     # Slate 900 (primary title banner)
C_HEADER_MID = "1E293B"      # Slate 800 (subtitle & table headers)
C_SECTION_BG = "E2E8F0"      # Slate 200 (section headers)
C_SECTION_FG = "1E293B"      # Slate 800

# Option semantic colors (shared across column headers & charts)
C_OPT_A = "475569"           # Slate 600 (Baseline)
C_OPT_B = "2563EB"           # Blue 600 (Mid-Tier)
C_OPT_C = "059669"           # Emerald 600 (Best / BIFL)
C_RENTAL = "D97706"          # Amber 600 (Rental / Alt)

# Input & Output tokens
C_INPUT_BG = "EFF6FF"        # Blue 50
C_INPUT_ALT = "E7F1FF"       # Slightly deeper Blue 50 for zebra rows
C_INPUT_FG = "1D4ED8"        # Blue 700
C_INPUT_BORDER = "BFDBFE"    # Blue 200

C_EMERALD_BG = "ECFDF5"      # Emerald 50
C_EMERALD_FG = "047857"      # Emerald 700
C_EMERALD_BORDER = "6EE7B7"  # Emerald 300

C_TEXT = "0F172A"            # Slate 900
C_TEXT_SUB = "475569"        # Slate 600
C_MUTED = "64748B"           # Slate 500
C_LIGHT_TEXT = "CBD5E1"      # Slate 300
C_BORDER = "E2E8F0"          # Slate 200 (hairline)
C_BORDER_CARD = "CBD5E1"     # Slate 300

FONT_FAMILY = "Roboto"

# Number formats (Google-Sheets-safe: no trailing dot on integers!)
FMT_CURRENCY = "$#,##0.00"
FMT_INT = "#,##0"
FMT_USES = '#,##0" uses"'
FMT_GENERAL = "General"
FMT_DEC1 = "0.0"
FMT_PCT = "0.0%"

# Reusable borders
S_HAIR = Side(border_style="thin", color=C_BORDER)
S_CARD = Side(border_style="thin", color=C_BORDER_CARD)
S_INPUT = Side(border_style="thin", color=C_INPUT_BORDER)
S_EMERALD_MED = Side(border_style="medium", color=C_EMERALD_BORDER)
S_DARK_MED = Side(border_style="medium", color=C_HEADER_MID)

B_CELL = Border(left=S_HAIR, right=S_HAIR, top=S_HAIR, bottom=S_HAIR)
B_INPUT = Border(left=S_INPUT, right=S_INPUT, top=S_INPUT, bottom=S_INPUT)
B_HEADER = Border(left=S_HAIR, right=S_HAIR, top=S_HAIR, bottom=S_DARK_MED)
B_HIGHLIGHT = Border(left=S_HAIR, right=S_HAIR, top=S_EMERALD_MED, bottom=S_EMERALD_MED)

# Reusable fills
FILL_CANVAS = PatternFill("solid", fgColor=C_CANVAS)
FILL_WHITE = PatternFill("solid", fgColor=C_WHITE)
FILL_ZEBRA = PatternFill("solid", fgColor=C_ZEBRA)
FILL_DARK = PatternFill("solid", fgColor=C_HEADER_DARK)
FILL_MID = PatternFill("solid", fgColor=C_HEADER_MID)
FILL_SECTION = PatternFill("solid", fgColor=C_SECTION_BG)
FILL_INPUT = PatternFill("solid", fgColor=C_INPUT_BG)
FILL_INPUT_ALT = PatternFill("solid", fgColor=C_INPUT_ALT)
FILL_EMERALD = PatternFill("solid", fgColor=C_EMERALD_BG)

# Reusable fonts
F_TITLE = Font(name=FONT_FAMILY, size=15, bold=True, color=C_WHITE)
F_SUBTITLE = Font(name=FONT_FAMILY, size=9.5, italic=True, color=C_LIGHT_TEXT)
F_SECTION = Font(name=FONT_FAMILY, size=9.5, bold=True, color=C_SECTION_FG)
F_TH = Font(name=FONT_FAMILY, size=9.5, bold=True, color=C_WHITE)
F_LABEL = Font(name=FONT_FAMILY, size=10, color=C_TEXT)
F_LABEL_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color=C_TEXT)
F_INPUT = Font(name=FONT_FAMILY, size=10, color=C_INPUT_FG)
F_INPUT_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color=C_INPUT_FG)
F_OUT = Font(name=FONT_FAMILY, size=10, color=C_TEXT)
F_OUT_HERO = Font(name=FONT_FAMILY, size=11.5, bold=True, color=C_EMERALD_FG)
F_NOTE = Font(name=FONT_FAMILY, size=9, italic=True, color=C_MUTED)
F_DATA_MUTED = Font(name=FONT_FAMILY, size=9, color=C_MUTED)
F_KPI_LABEL = Font(name=FONT_FAMILY, size=8.5, bold=True, color=C_MUTED)
F_KPI_VAL = Font(name=FONT_FAMILY, size=14, bold=True, color=C_TEXT)
F_KPI_VAL_EMERALD = Font(name=FONT_FAMILY, size=15, bold=True, color=C_EMERALD_FG)
F_KPI_SUB = Font(name=FONT_FAMILY, size=8.5, color=C_MUTED)


def fill_canvas(ws, max_row: int, max_col: int):
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).fill = FILL_CANVAS


def style_range(ws, cell_range, font=None, fill=None, border=None, alignment=None, num_format=None):
    for row in ws[cell_range]:
        for cell in row:
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if border:
                cell.border = border
            if alignment:
                cell.alignment = alignment
            if num_format:
                cell.number_format = num_format


def draw_kpi_card(
    ws,
    start_col: int,
    end_col: int,
    top_row: int,
    accent_color: str,
    title: str,
    formula: str,
    subtitle: str,
    num_format: str | None = None,
    val_font: Font = F_KPI_VAL,
    include_sub: bool = True,
):
    """Draw a clean elevated KPI card with a colored top accent bar."""
    bottom_row = top_row + (2 if include_sub else 1)
    top_accent = Side(border_style="thick", color=accent_color)

    ws.merge_cells(start_row=top_row, start_column=start_col, end_row=top_row, end_column=end_col)
    ws.merge_cells(start_row=top_row + 1, start_column=start_col, end_row=top_row + 1, end_column=end_col)
    if include_sub:
        ws.merge_cells(start_row=bottom_row, start_column=start_col, end_row=bottom_row, end_column=end_col)

    for r in range(top_row, bottom_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = FILL_WHITE
            cell.border = Border(
                top=top_accent if r == top_row else None,
                bottom=S_CARD if r == bottom_row else None,
                left=S_CARD if c == start_col else None,
                right=S_CARD if c == end_col else None,
            )

    lbl = ws.cell(row=top_row, column=start_col, value=title)
    lbl.font = F_KPI_LABEL
    lbl.alignment = Alignment(horizontal="center", vertical="bottom")

    val = ws.cell(row=top_row + 1, column=start_col, value=formula)
    val.font = val_font
    val.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if num_format:
        val.number_format = num_format

    if include_sub:
        sub = ws.cell(row=bottom_row, column=start_col, value=subtitle)
        sub.font = F_KPI_SUB
        sub.alignment = Alignment(horizontal="center", vertical="top")


def make_chart_font(size_pt: int = 9, bold: bool = False, color: str = C_TEXT_SUB) -> RichText:
    cp = CharacterProperties(sz=size_pt * 100, b=bold, solidFill=color)
    pp = ParagraphProperties(defRPr=cp)
    return RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])


def style_chart_common(chart, title: str, y_title: str | None = None, x_title: str | None = None, y_fmt: str = "$#,##0.00"):
    chart.title = title
    chart.style = 10
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "nextTo"
    chart.y_axis.number_format = y_fmt
    chart.y_axis.scaling.min = 0
    chart.y_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=C_BORDER, w=6350))
    if y_title:
        chart.y_axis.title = y_title
    if x_title:
        chart.x_axis.title = x_title
    # Clean white chart area with subtle slate border
    chart.graphical_properties = GraphicalProperties(
        solidFill=C_WHITE,
        ln=LineProperties(solidFill=C_BORDER_CARD, w=9525),
    )
    chart.plot_area.graphicalProperties = GraphicalProperties(
        solidFill=C_WHITE,
        ln=LineProperties(noFill=True),
    )
    chart.x_axis.txPr = make_chart_font(9, False, C_TEXT_SUB)
    chart.y_axis.txPr = make_chart_font(9, False, C_TEXT_SUB)


def build_quick_calc_tab(wb: Workbook):
    """Minimalistic single-item Cost Per Use calculator (pure typography & whitespace)."""
    ws = wb.active
    ws.title = "Quick Calc"
    ws.views.sheetView[0].showGridLines = False
    ws.freeze_panes = None

    fill_canvas(ws, max_row=28, max_col=9)

    # Outer canvas margins + inner card padding (B and F) + 3 balanced content columns (C, D, E)
    ws.row_dimensions[1].height = 22
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 2.8
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 25
    ws.column_dimensions["F"].width = 2.8
    ws.column_dimensions["G"].width = 4

    # Card surface (B2:F18): crisp white container with subtle hairline perimeter
    for r in range(2, 19):
        for c in range(2, 7):
            cell = ws.cell(row=r, column=c)
            cell.fill = FILL_WHITE
            cell.border = Border(
                top=S_CARD if r == 2 else None,
                bottom=S_CARD if r == 18 else None,
                left=S_CARD if c == 2 else None,
                right=S_CARD if c == 6 else None,
            )

    ws.row_dimensions[2].height = 14

    # Header (Rows 3-4)
    ws.merge_cells("C3:E3")
    ws["C3"] = "Cost Per Use"
    ws["C3"].font = Font(name=FONT_FAMILY, size=17, bold=True, color=C_TEXT)
    ws["C3"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 28

    ws.merge_cells("C4:E4")
    ws["C4"] = "Minimal single-item calculator — edit the 5 blue cells."
    ws["C4"].font = Font(name=FONT_FAMILY, size=9.5, color=C_MUTED)
    ws["C4"].alignment = Alignment(horizontal="left", vertical="top")
    for col in ["C", "D", "E"]:
        ws[f"{col}4"].border = Border(bottom=S_HAIR)
    ws.row_dimensions[4].height = 22

    ws.row_dimensions[5].height = 10

    # 5 Minimal Inputs (Rows 6-10)
    inputs = [
        (6, "Item", "Darn Tough Wool Socks", None, True),
        (7, "Purchase Price", 20, FMT_CURRENCY, False),
        (8, "Usage Frequency", 50, FMT_GENERAL, False),
        (9, "Frequency Unit", "Year", None, False),
        (10, "Estimated Life (Years)", 4, FMT_GENERAL, False),
    ]
    for r, label, default_val, num_fmt, is_text_bold in inputs:
        ws.row_dimensions[r].height = 30
        ws.merge_cells(f"C{r}:D{r}")
        ws[f"C{r}"] = label
        ws[f"C{r}"].font = Font(name=FONT_FAMILY, size=10.5, color=C_TEXT_SUB)
        ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="center")
        ws[f"C{r}"].border = Border(bottom=S_HAIR)
        ws[f"D{r}"].border = Border(bottom=S_HAIR)

        val_cell = ws[f"E{r}"]
        val_cell.value = default_val
        val_cell.font = Font(
            name=FONT_FAMILY,
            size=10.5,
            bold=is_text_bold,
            color=C_INPUT_FG,
        )
        val_cell.fill = FILL_INPUT
        val_cell.border = B_INPUT
        val_cell.alignment = Alignment(
            horizontal="right" if isinstance(default_val, (int, float)) or r == 6 else "center",
            vertical="center",
            wrap_text=(r == 6),
        )
        if num_fmt:
            val_cell.number_format = num_fmt

    dv_unit = DataValidation(type="list", formula1=f'"{",".join(FREQ_UNITS)}"', allow_blank=True)
    ws.add_data_validation(dv_unit)
    dv_unit.add("E9")

    ws.row_dimensions[11].height = 16

    # Inset Hero Output Block (C12:E14)
    emerald_edge = Side(border_style="thin", color=C_EMERALD_BORDER)
    ws.merge_cells("C12:E12")
    ws.merge_cells("C13:E13")
    ws.merge_cells("C14:E14")
    for r in range(12, 15):
        for c in range(3, 6):
            cell = ws.cell(row=r, column=c)
            cell.fill = FILL_EMERALD
            cell.border = Border(
                top=emerald_edge if r == 12 else None,
                bottom=emerald_edge if r == 14 else None,
                left=emerald_edge if c == 3 else None,
                right=emerald_edge if c == 5 else None,
            )

    ws.row_dimensions[12].height = 22
    ws["C12"] = "COST PER USE"
    ws["C12"].font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=C_EMERALD_FG)
    ws["C12"].alignment = Alignment(horizontal="center", vertical="bottom")

    ws.row_dimensions[13].height = 46
    ws["C13"] = '=IF(OR(E7="",C17="",C17=0),"—",E7/C17)'
    ws["C13"].font = Font(name=FONT_FAMILY, size=26, bold=True, color=C_EMERALD_FG)
    ws["C13"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C13"].number_format = FMT_CURRENCY

    ws.row_dimensions[14].height = 22
    ws["C14"] = '=IF(OR(C17="",C17=0),"",IF(E6="","over "&TEXT(C17,"#,##0")&" estimated uses","for "&E6&" over "&TEXT(C17,"#,##0")&" uses"))'
    ws["C14"].font = Font(name=FONT_FAMILY, size=9, color="065F46")
    ws["C14"].alignment = Alignment(horizontal="center", vertical="top")

    ws.row_dimensions[15].height = 14

    # 3 Quiet Supporting Metrics (Rows 16-17)
    ws.row_dimensions[16].height = 22
    ws.row_dimensions[17].height = 28

    metrics = [
        ("C", "TOTAL USES", '=IF(OR(E8="",E9="",E10=""),"",IFERROR(IF(E9="Day",E8*365,IF(E9="Week",E8*52,IF(E9="Month",E8*12,IF(E9="Year",E8,""))))*E10,""))', FMT_INT),
        ("D", "COST / YEAR", '=IF(OR(E7="",E10="",E10=0),"",E7/E10)', FMT_CURRENCY),
        ("E", "COST / MONTH", '=IF(D17="","",D17/12)', FMT_CURRENCY),
    ]
    for col, title, formula, num_fmt in metrics:
        lbl = ws[f"{col}16"]
        lbl.value = title
        lbl.font = F_KPI_LABEL
        lbl.alignment = Alignment(horizontal="center", vertical="bottom")
        lbl.border = Border(top=S_HAIR)

        val = ws[f"{col}17"]
        val.value = formula
        val.font = Font(name=FONT_FAMILY, size=13, bold=True, color=C_TEXT)
        val.alignment = Alignment(horizontal="center", vertical="center")
        val.number_format = num_fmt

    ws.row_dimensions[18].height = 14

    # Subtle canvas footer
    ws.row_dimensions[19].height = 24
    ws.merge_cells("B19:F19")
    ws["B19"] = "Need resale, ongoing costs, or A/B comparison? Switch to Calculator & Compare."
    ws["B19"].font = Font(name=FONT_FAMILY, size=8.5, italic=True, color=C_MUTED)
    ws["B19"].alignment = Alignment(horizontal="center", vertical="center")


def build_calculator_tab(wb: Workbook):
    ws = wb.create_sheet(title="Calculator & Compare")
    ws.views.sheetView[0].showGridLines = False
    ws.freeze_panes = None  # No freeze lines cutting through banners/cards

    fill_canvas(ws, max_row=60, max_col=15)

    # Outer & inner margins
    ws.row_dimensions[1].height = 12
    ws.column_dimensions["A"].width = 2.5
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 19
    ws.column_dimensions["D"].width = 19
    ws.column_dimensions["E"].width = 21
    ws.column_dimensions["F"].width = 34
    ws.column_dimensions["G"].width = 2.5
    for col in ["H", "I", "J", "K", "L", "M"]:
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["N"].width = 2.5

    # Title Banner (B2:M3)
    ws.merge_cells("B2:M2")
    ws["B2"] = "Cost Per Use — Purchase Calculator & Comparison"
    style_range(
        ws, "B2:M2", font=F_TITLE, fill=FILL_DARK,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[2].height = 36

    ws.merge_cells("B3:M3")
    ws["B3"] = "Compare up to 3 options side-by-side (Budget vs. Mid-Tier vs. Buy-It-For-Life) and visualize break-even over time. Edit blue cells."
    style_range(
        ws, "B3:M3", font=F_SUBTITLE, fill=FILL_MID,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 12

    # Executive KPI Cards (Rows 5-7)
    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 32
    ws.row_dimensions[7].height = 18

    draw_kpi_card(
        ws, 2, 3, 5, C_OPT_C,
        "BEST VALUE OPTION",
        '=IF(COUNT(C26:E26)=0,"—",INDEX(C11:E11,1,MATCH(MIN(C26:E26),C26:E26,0)))',
        "Lowest cost per use",
        val_font=Font(name=FONT_FAMILY, size=12, bold=True, color=C_TEXT),
    )
    draw_kpi_card(
        ws, 4, 5, 5, C_OPT_C,
        "WINNING COST / USE",
        '=IF(COUNT(C26:E26)=0,"—",MIN(C26:E26))',
        "Amortized over full life",
        num_format=FMT_CURRENCY,
        val_font=F_KPI_VAL_EMERALD,
    )
    draw_kpi_card(
        ws, 6, 6, 5, C_OPT_B,
        "SAVINGS / 100 USES (C vs A)",
        '=IF(OR(C26="",E26=""),"—",(C26-E26)*100)',
        "Net savings upgrading to C",
        num_format=FMT_CURRENCY,
        val_font=F_KPI_VAL,
    )
    draw_kpi_card(
        ws, 8, 10, 5, C_OPT_A,
        "BREAK-EVEN (OPTION C vs A)",
        '=IF(OR(C26="",E26=""),"—",IF(E26>=C26,"Not cheaper",ROUNDUP(E25/C26,0)))',
        "Uses until C beats replacing A",
        num_format=FMT_USES,
        val_font=F_KPI_VAL,
    )
    draw_kpi_card(
        ws, 11, 13, 5, C_RENTAL,
        "BREAK-EVEN vs. RENTAL (BEST)",
        '=IF(OR(C34="",C34=0,COUNT(C26:E26)=0),"—",IF(MIN(C26:E26)>=C34,"Rent wins",ROUNDUP(INDEX(C25:E25,1,MATCH(MIN(C26:E26),C26:E26,0))/C34,0)))',
        "Uses for best option to beat renting",
        num_format=FMT_USES,
        val_font=F_KPI_VAL,
    )

    ws.row_dimensions[8].height = 12

    # Calculator Column Headers (Row 9) — semantic header colors matching chart series
    header_specs = [
        ("B9", "Parameter", C_HEADER_MID, "left"),
        ("C9", "Option A (Baseline)", C_OPT_A, "center"),
        ("D9", "Option B (Mid-Tier)", C_OPT_B, "center"),
        ("E9", "Option C (Best / BIFL)", C_OPT_C, "center"),
        ("F9", "Notes / Guidance", C_HEADER_MID, "left"),
    ]
    ws.row_dimensions[9].height = 30
    for coord, text, bg, align_h in header_specs:
        cell = ws[coord]
        cell.value = text
        cell.font = F_TH
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.border = B_HEADER
        cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0, wrap_text=True)

    # Section 1: Inputs (Row 10 banner, Rows 11-19 inputs)
    ws.merge_cells("B10:F10")
    ws["B10"] = "1. ITEM PARAMETERS  (EDIT BLUE CELLS)"
    style_range(
        ws, "B10:F10", font=F_SECTION, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[10].height = 24

    input_rows = [
        # (row, label, sample A, sample B, sample C, num_fmt, note, height, bold_input)
        (11, "Item / Option Name", "Poor Socks", "Good Socks", "Darn Tough Wool Socks", None, "Name of item or quality tier", 30, True),
        (12, "Category", "Apparel", "Apparel", "Apparel", None, "Select category from dropdown", 24, False),
        (13, "Purchase Price ($)", 5, 10, 20, FMT_CURRENCY, "Upfront cost (enter $0 for subscriptions)", 24, False),
        (14, "Annual Ongoing / Maintenance ($/yr)", 0, 0, 0, FMT_CURRENCY, "Yearly fees/upkeep (monthly sub: =Monthly*12)", 24, False),
        (15, "Estimated Resale Value ($)", 0, 0, 0, FMT_CURRENCY, "Expected value recovered if resold", 24, False),
        (16, "Usage Frequency (Count)", 10, 50, 50, FMT_GENERAL, "How many times used per unit below", 24, False),
        (17, "Frequency Unit", "Year", "Year", "Year", None, "Day, Week, Month, or Year", 24, False),
        (18, "Estimated Product Life (Years)", 1, 1, 4, FMT_GENERAL, "Expected useful lifespan in years", 24, False),
        (19, "Durability / Value Rating (1-5)", 1, 3, 5, FMT_INT, "1 = disposable, 5 = repairable / BIFL", 24, False),
    ]

    for row_idx, label, v_a, v_b, v_c, num_fmt, note, row_h, is_bold in input_rows:
        ws.row_dimensions[row_idx].height = row_h
        c_lbl = ws[f"B{row_idx}"]
        c_lbl.value = label
        c_lbl.font = F_LABEL_BOLD if is_bold else F_LABEL
        c_lbl.fill = FILL_WHITE
        c_lbl.border = B_CELL
        c_lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        for col_letter, val in [("C", v_a), ("D", v_b), ("E", v_c)]:
            cell = ws[f"{col_letter}{row_idx}"]
            cell.value = val
            cell.font = F_INPUT_BOLD if is_bold else F_INPUT
            cell.fill = FILL_INPUT
            cell.border = B_INPUT
            cell.alignment = Alignment(
                horizontal="right" if isinstance(val, (int, float)) else "center",
                vertical="center",
                wrap_text=is_bold,
            )
            if num_fmt:
                cell.number_format = num_fmt

        c_note = ws[f"F{row_idx}"]
        c_note.value = note
        c_note.font = F_NOTE
        c_note.fill = FILL_WHITE
        c_note.border = B_CELL
        c_note.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Dropdown validations for Calculator
    dv_cat = DataValidation(type="list", formula1=f'"{",".join(CATEGORIES)}"', allow_blank=True)
    dv_unit = DataValidation(type="list", formula1=f'"{",".join(FREQ_UNITS)}"', allow_blank=True)
    dv_rate = DataValidation(type="list", formula1=f'"{",".join(RATINGS)}"', allow_blank=True)
    ws.add_data_validation(dv_cat)
    ws.add_data_validation(dv_unit)
    ws.add_data_validation(dv_rate)
    dv_cat.add("C12:E12")
    dv_unit.add("C17:E17")
    dv_rate.add("C19:E19")

    ws.row_dimensions[20].height = 10

    # Section 2: Outputs (Row 21 banner, Rows 22-28 outputs)
    ws.merge_cells("B21:F21")
    ws["B21"] = "2. COST PER USE OUTPUTS  (AUTO-CALCULATED)"
    style_range(
        ws, "B21:F21", font=F_SECTION, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[21].height = 24

    output_specs = [
        (22, "Uses per Year", lambda c: f'=IF(OR({c}16="",{c}17=""),"",IF({c}17="Day",{c}16*365,IF({c}17="Week",{c}16*52,IF({c}17="Month",{c}16*12,IF({c}17="Year",{c}16,"")))))', FMT_INT, False, "Annualized usage count"),
        (23, "Total Estimated Uses", lambda c: f'=IF(OR({c}22="",{c}18=""),"",{c}22*{c}18)', FMT_INT, False, "Uses/Year x Estimated Life"),
        (24, "Total Ongoing Cost ($)", lambda c: f'=IF(OR({c}14="",{c}18=""),"",{c}14*{c}18)', FMT_CURRENCY, False, "Annual Ongoing x Estimated Life"),
        (25, "Total Net Lifecycle Cost ($)", lambda c: f'=IF({c}13="","",{c}13+IF({c}24="",0,{c}24)-IF({c}15="",0,{c}15))', FMT_CURRENCY, False, "Purchase Price + Ongoing - Resale"),
        (26, "Cost per Use ($)", lambda c: f'=IF(OR({c}25="",{c}23="",{c}23=0),"",{c}25/{c}23)', FMT_CURRENCY, True, "Primary metric: Total Net Cost / Total Uses"),
        (27, "Annualized Cost ($/yr)", lambda c: f'=IF(OR({c}25="",{c}18="",{c}18=0),"",{c}25/{c}18)', FMT_CURRENCY, False, "Total Net Cost / Estimated Life"),
        (28, "Monthly Cost ($/mo)", lambda c: f'=IF({c}27="","",{c}27/12)', FMT_CURRENCY, False, "Annualized Cost / 12"),
    ]

    for row_idx, label, formula_fn, num_fmt, is_hero, note in output_specs:
        ws.row_dimensions[row_idx].height = 29 if is_hero else 24
        row_fill = FILL_EMERALD if is_hero else FILL_WHITE
        row_border = B_HIGHLIGHT if is_hero else B_CELL

        c_lbl = ws[f"B{row_idx}"]
        c_lbl.value = label
        c_lbl.font = F_LABEL_BOLD if is_hero else F_LABEL
        c_lbl.fill = row_fill
        c_lbl.border = row_border
        c_lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        for col_letter in ["C", "D", "E"]:
            cell = ws[f"{col_letter}{row_idx}"]
            cell.value = formula_fn(col_letter)
            cell.font = F_OUT_HERO if is_hero else F_OUT
            cell.fill = row_fill
            cell.border = row_border
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = num_fmt

        c_note = ws[f"F{row_idx}"]
        c_note.value = note
        c_note.font = F_NOTE
        c_note.fill = row_fill
        c_note.border = row_border
        c_note.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.row_dimensions[29].height = 10

    # Section 3: Decision & Break-Even Analysis (Rows 30-35)
    ws.merge_cells("B30:F30")
    ws["B30"] = "3. DECISION & BREAK-EVEN ANALYSIS"
    style_range(
        ws, "B30:F30", font=F_SECTION, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[30].height = 24

    # Row 31: Winner (merged C31:D31 so long names never clip!)
    ws.row_dimensions[31].height = 28
    ws["B31"] = "Lowest Cost per Use Winner"
    ws["B31"].font = F_LABEL_BOLD
    ws["B31"].fill = FILL_WHITE
    ws["B31"].border = B_CELL
    ws["B31"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.merge_cells("C31:D31")
    ws["C31"] = '=IF(COUNT(C26:E26)=0,"",INDEX(C11:E11,1,MATCH(MIN(C26:E26),C26:E26,0)))'
    style_range(
        ws, "C31:D31", font=F_OUT_HERO, fill=FILL_EMERALD, border=B_HIGHLIGHT,
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
    )

    ws["E31"] = '=IF(COUNT(C26:E26)=0,"",MIN(C26:E26))'
    ws["E31"].font = F_OUT_HERO
    ws["E31"].fill = FILL_EMERALD
    ws["E31"].border = B_HIGHLIGHT
    ws["E31"].alignment = Alignment(horizontal="right", vertical="center")
    ws["E31"].number_format = FMT_CURRENCY

    ws["F31"] = "Best long-term value based purely on Cost/Use"
    ws["F31"].font = F_NOTE
    ws["F31"].fill = FILL_WHITE
    ws["F31"].border = B_CELL
    ws["F31"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Row 32: Break-Even Uses vs. Option A
    ws.row_dimensions[32].height = 25
    ws["B32"] = "Break-Even Uses vs. Option A"
    ws["B32"].font = F_LABEL
    ws["B32"].fill = FILL_WHITE
    ws["B32"].border = B_CELL
    ws["B32"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws["C32"] = "Baseline"
    ws["C32"].font = F_NOTE
    ws["C32"].fill = FILL_ZEBRA
    ws["C32"].border = B_CELL
    ws["C32"].alignment = Alignment(horizontal="center", vertical="center")

    for col_letter in ["D", "E"]:
        cell = ws[f"{col_letter}32"]
        cell.value = f'=IF(OR($C$26="",{col_letter}26=""),"",IF({col_letter}26>=$C$26,"Not cheaper/use",ROUNDUP({col_letter}25/$C$26,0)))'
        cell.font = F_OUT
        cell.fill = FILL_WHITE
        cell.border = B_CELL
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = FMT_INT

    ws["F32"] = "Uses until upgrading beats replacing Option A"
    ws["F32"].font = F_NOTE
    ws["F32"].fill = FILL_WHITE
    ws["F32"].border = B_CELL
    ws["F32"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Row 33: Savings per 100 Uses vs Option A
    ws.row_dimensions[33].height = 25
    ws["B33"] = "Savings per 100 Uses vs. Option A ($)"
    ws["B33"].font = F_LABEL
    ws["B33"].fill = FILL_WHITE
    ws["B33"].border = B_CELL
    ws["B33"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws["C33"] = 0
    ws["C33"].font = F_NOTE
    ws["C33"].fill = FILL_ZEBRA
    ws["C33"].border = B_CELL
    ws["C33"].alignment = Alignment(horizontal="right", vertical="center")
    ws["C33"].number_format = FMT_CURRENCY

    for col_letter in ["D", "E"]:
        cell = ws[f"{col_letter}33"]
        cell.value = f'=IF(OR($C$26="",{col_letter}26=""),"",($C$26-{col_letter}26)*100)'
        cell.font = F_OUT
        cell.fill = FILL_WHITE
        cell.border = B_CELL
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = FMT_CURRENCY

    ws["F33"] = "Positive = net savings per 100 uses vs. Option A"
    ws["F33"].font = F_NOTE
    ws["F33"].fill = FILL_WHITE
    ws["F33"].border = B_CELL
    ws["F33"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Row 34: Rental / Pay-Per-Use Alternative Input
    ws.row_dimensions[34].height = 25
    ws["B34"] = "Rental / Pay-Per-Use Alternative ($/use)"
    ws["B34"].font = F_LABEL
    ws["B34"].fill = FILL_WHITE
    ws["B34"].border = B_CELL
    ws["B34"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws["C34"] = 1.00
    ws["C34"].font = F_INPUT_BOLD
    ws["C34"].fill = FILL_INPUT
    ws["C34"].border = B_INPUT
    ws["C34"].alignment = Alignment(horizontal="right", vertical="center")
    ws["C34"].number_format = FMT_CURRENCY

    ws.merge_cells("D34:E34")
    ws["D34"] = "← Enter rental or single-use price"
    style_range(
        ws, "D34:E34", font=F_NOTE, fill=FILL_WHITE, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )

    ws["F34"] = "Cost to rent, borrow, or use disposable per use"
    ws["F34"].font = F_NOTE
    ws["F34"].fill = FILL_WHITE
    ws["F34"].border = B_CELL
    ws["F34"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Row 35: Break-Even Uses vs. Renting
    ws.row_dimensions[35].height = 25
    ws["B35"] = "Break-Even Uses vs. Renting"
    ws["B35"].font = F_LABEL
    ws["B35"].fill = FILL_WHITE
    ws["B35"].border = B_CELL
    ws["B35"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for col_letter in ["C", "D", "E"]:
        cell = ws[f"{col_letter}35"]
        cell.value = f'=IF(OR($C$34="",$C$34=0,{col_letter}25=""),"",IF({col_letter}26>=$C$34,"Rent is cheaper",ROUNDUP({col_letter}25/$C$34,0)))'
        cell.font = F_OUT
        cell.fill = FILL_WHITE
        cell.border = B_CELL
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = FMT_INT

    ws["F35"] = "Uses needed for ownership to beat renting"
    ws["F35"].font = F_NOTE
    ws["F35"].fill = FILL_WHITE
    ws["F35"].border = B_CELL
    ws["F35"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.row_dimensions[36].height = 10

    # Section 4: Qualitative Buy Wisely Prompts (Rows 37-41)
    ws.merge_cells("B37:F37")
    ws["B37"] = "4. BUY WISELY DURABILITY CHECKS  (BEFORE YOU BUY)"
    style_range(
        ws, "B37:F37", font=F_SECTION, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[37].height = 24

    prompts = [
        "1. Longevity & Materials — Will it stay useful for 5+ years, age well, and avoid planned obsolescence?",
        "2. Repair & Parts — Can it be disassembled and repaired with standard, non-proprietary parts or batteries?",
        "3. Resale & Guarantee — Does it carry a strong warranty, and would someone buy it second-hand?",
        "4. Frequency Reality Check — If new to a hobby, borrow or rent first; the lowest $/use is skipping what you won't use.",
    ]
    for offset, text in enumerate(prompts):
        r = 38 + offset
        ws.merge_cells(f"B{r}:F{r}")
        ws[f"B{r}"] = text
        style_range(
            ws, f"B{r}:F{r}", font=F_NOTE, fill=FILL_WHITE, border=B_CELL,
            alignment=Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True),
        )
        ws.row_dimensions[r].height = 24

    ws.row_dimensions[42].height = 18

    # =========================================================================
    # AUTO-GENERATED CHART DATA TABLES (Rows 43-55, cleanly styled & local)
    # =========================================================================
    ws.merge_cells("B43:E43")
    ws["B43"] = "AUTO-GENERATED CURVE DATA (CUMULATIVE REPLACEMENT COST)"
    style_range(
        ws, "B43:E43", font=F_KPI_LABEL, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[43].height = 22

    # Curve table headers (Row 44)
    ws.row_dimensions[44].height = 22
    ws["B44"] = "Uses Milestone"
    ws["C44"] = '=IF(C11="","Option A",C11)'
    ws["D44"] = '=IF(D11="","Option B",D11)'
    ws["E44"] = '=IF(E11="","Option C",E11)'
    style_range(
        ws, "B44:E44", font=F_KPI_LABEL, fill=FILL_WHITE, border=B_CELL,
        alignment=Alignment(horizontal="center", vertical="center"),
    )

    # 10 dynamic milestones from 10% to 100% of MAX(Total Uses)
    for idx in range(1, 11):
        r = 44 + idx
        ws.row_dimensions[r].height = 20
        ws[f"B{r}"] = f'=IF(MAX($C$23:$E$23)=0,"",ROUNDUP(MAX($C$23:$E$23)*{idx}/10,0))'
        ws[f"B{r}"].number_format = FMT_INT
        ws[f"B{r}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"B{r}"].font = F_DATA_MUTED
        ws[f"B{r}"].fill = FILL_WHITE
        ws[f"B{r}"].border = B_CELL

        for col_letter in ["C", "D", "E"]:
            cell = ws[f"{col_letter}{r}"]
            cell.value = f'=IF(OR($B{r}="",${col_letter}$23="",${col_letter}$23=0,${col_letter}$25=""),"",ROUNDUP($B{r}/${col_letter}$23,0)*${col_letter}$25)'
            cell.number_format = FMT_CURRENCY
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font = F_DATA_MUTED
            cell.fill = FILL_WHITE
            cell.border = B_CELL

    # Bar chart backing table (H43:I48)
    ws.merge_cells("H43:I43")
    ws["H43"] = "CHART DATA: COST / USE"
    style_range(
        ws, "H43:I43", font=F_KPI_LABEL, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws["H44"] = "Option"
    ws["I44"] = "Cost per Use ($)"
    style_range(
        ws, "H44:I44", font=F_KPI_LABEL, fill=FILL_WHITE, border=B_CELL,
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    bar_rows = [
        (45, '=IF(C11="","Option A",C11)', '=IF(C26="",0,C26)'),
        (46, '=IF(D11="","Option B",D11)', '=IF(D26="",0,D26)'),
        (47, '=IF(E11="","Option C",E11)', '=IF(E26="",0,E26)'),
        (48, "Rental / Alt", '=IF(C34="",0,C34)'),
    ]
    for r, lbl_f, val_f in bar_rows:
        ws[f"H{r}"] = lbl_f
        ws[f"H{r}"].font = F_DATA_MUTED
        ws[f"H{r}"].fill = FILL_WHITE
        ws[f"H{r}"].border = B_CELL
        ws[f"H{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

        ws[f"I{r}"] = val_f
        ws[f"I{r}"].font = F_DATA_MUTED
        ws[f"I{r}"].fill = FILL_WHITE
        ws[f"I{r}"].border = B_CELL
        ws[f"I{r}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"I{r}"].number_format = FMT_CURRENCY

    # =========================================================================
    # VISUALIZATION PANEL (Right Side: H9:M41)
    # =========================================================================
    # Chart 1: Cost per Use Comparison (Column Chart)
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.overlap = 0
    style_chart_common(bar, "Cost per Use Comparison ($)", y_title="$ / Use", y_fmt="$#,##0.00")
    bar.legend = None
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True
    bar.width = 16.4
    bar.height = 13.2

    bar_data = Reference(ws, min_col=9, min_row=44, max_row=48)
    bar_cats = Reference(ws, min_col=8, min_row=45, max_row=48)
    bar.add_data(bar_data, titles_from_data=True)
    bar.set_categories(bar_cats)

    # Color each bar semantically (Slate, Blue, Emerald, Amber)
    bar_colors = [C_OPT_A, C_OPT_B, C_OPT_C, C_RENTAL]
    bar.series[0].graphicalProperties.solidFill = C_OPT_B
    bar.series[0].graphicalProperties.line.noFill = True
    for i, hex_color in enumerate(bar_colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = hex_color
        pt.graphicalProperties.line.noFill = True
        bar.series[0].data_points.append(pt)

    ws.add_chart(bar, "H9")

    # Chart 2: Cumulative Replacement Cost Over Uses (Line Chart)
    line = LineChart()
    line.grouping = "standard"
    style_chart_common(
        line,
        "Cumulative Cost Over Uses (Replacement Break-Even)",
        y_title="Total Spend ($)",
        x_title="Number of Uses",
        y_fmt="$#,##0",
    )
    line.legend.position = "b"
    line.width = 16.4
    line.height = 12.8

    line_data = Reference(ws, min_col=3, max_col=5, min_row=44, max_row=54)
    line_cats = Reference(ws, min_col=2, min_row=45, max_row=54)
    line.add_data(line_data, titles_from_data=True)
    line.set_categories(line_cats)

    line_styles = [
        (C_OPT_A, 28575, "sysDash"),   # 2.25 pt dashed slate for Option A
        (C_OPT_B, 28575, "solid"),     # 2.25 pt solid blue for Option B
        (C_OPT_C, 34925, "solid"),     # 2.75 pt solid emerald for Option C
    ]
    for s, (hex_color, width_emu, dash) in zip(line.series, line_styles):
        s.smooth = False
        s.graphicalProperties.line = LineProperties(solidFill=hex_color, w=width_emu)
        if dash != "solid":
            s.graphicalProperties.line.prstDash = dash

    ws.add_chart(line, "H26")


def build_master_tracker_tab(wb: Workbook):
    ws = wb.create_sheet(title="Master Tracker")
    ws.views.sheetView[0].showGridLines = False
    ws.freeze_panes = "A9"  # Horizontal freeze ONLY (Rows 1-8) — no vertical slice!

    max_row = 108  # 100 data rows: Rows 9..108
    fill_canvas(ws, max_row=max_row + 4, max_col=20)

    ws.row_dimensions[1].height = 10
    ws.column_dimensions["A"].width = 2.5
    ws.column_dimensions["T"].width = 2.5

    # Title Banner (B2:S3)
    ws.merge_cells("B2:S2")
    ws["B2"] = "Master Purchase & Cost-Per-Use Tracker"
    style_range(
        ws, "B2:S2", font=F_TITLE, fill=FILL_DARK,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[2].height = 36

    ws.merge_cells("B3:S3")
    ws["B3"] = "Track items you own or are considering. Blue columns are inputs; white & emerald columns auto-calculate. Sample rows can be overwritten."
    style_range(
        ws, "B3:S3", font=F_SUBTITLE, fill=FILL_MID,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 10

    # Top KPI Strip (Rows 5-6)
    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 30

    draw_kpi_card(ws, 2, 3, 5, C_OPT_A, "TRACKED ITEMS", "=COUNTA(B9:B108)", "", FMT_INT, F_KPI_VAL, include_sub=False)
    draw_kpi_card(ws, 4, 5, 5, C_OPT_C, "OWNED ITEMS", '=COUNTIF(D9:D108,"Owned")', "", FMT_INT, F_KPI_VAL, include_sub=False)
    draw_kpi_card(ws, 6, 7, 5, C_RENTAL, "CONSIDERING", '=COUNTIF(D9:D108,"Considering")', "", FMT_INT, F_KPI_VAL, include_sub=False)
    draw_kpi_card(ws, 8, 10, 5, C_OPT_B, "TOTAL NET SPEND (OWNED)", '=SUMIF(D9:D108,"Owned",N9:N108)', "", FMT_CURRENCY, F_KPI_VAL, include_sub=False)
    draw_kpi_card(ws, 11, 13, 5, C_OPT_B, "ANNUALIZED COST (OWNED)", '=SUMIF(D9:D108,"Owned",P9:P108)', "", FMT_CURRENCY, F_KPI_VAL, include_sub=False)
    draw_kpi_card(ws, 14, 16, 5, C_OPT_C, "AVG COST / USE (ALL)", '=IFERROR(AVERAGE(O9:O108),"")', "", FMT_CURRENCY, F_KPI_VAL_EMERALD, include_sub=False)
    draw_kpi_card(ws, 17, 19, 5, C_OPT_A, "AVG DURABILITY (1-5)", '=IFERROR(AVERAGE(K9:K108),"")', "", FMT_DEC1, F_KPI_VAL, include_sub=False)

    ws.row_dimensions[7].height = 10

    # Table Columns (B..S)
    columns = [
        ("B", "Item Name", 30, True, None, "left"),
        ("C", "Category", 21, True, None, "center"),
        ("D", "Status", 14, True, None, "center"),
        ("E", "Purchase Price ($)", 15, True, FMT_CURRENCY, "right"),
        ("F", "Annual Ongoing ($/yr)", 16, True, FMT_CURRENCY, "right"),
        ("G", "Est. Resale ($)", 14, True, FMT_CURRENCY, "right"),
        ("H", "Usage Freq", 12, True, FMT_GENERAL, "right"),
        ("I", "Freq Unit", 11, True, None, "center"),
        ("J", "Est. Life (Yrs)", 13, True, FMT_GENERAL, "right"),
        ("K", "Durability (1-5)", 14, True, FMT_INT, "center"),
        ("L", "Uses / Year", 13, False, FMT_INT, "right"),
        ("M", "Total Uses", 13, False, FMT_INT, "right"),
        ("N", "Total Net Cost ($)", 16, False, FMT_CURRENCY, "right"),
        ("O", "Cost / Use ($)", 14, False, FMT_CURRENCY, "right"),
        ("P", "Cost / Year ($)", 15, False, FMT_CURRENCY, "right"),
        ("Q", "Cost / Month ($)", 15, False, FMT_CURRENCY, "right"),
        ("R", "Value Tier", 15, False, None, "center"),
        ("S", "Notes (Warranty / Repairability)", 34, True, None, "left"),
    ]

    ws.row_dimensions[8].height = 32
    for col_letter, title, width, is_input, _, _ in columns:
        cell = ws[f"{col_letter}8"]
        cell.value = title
        cell.font = F_TH
        cell.fill = PatternFill("solid", fgColor=C_OPT_C if col_letter == "O" else C_HEADER_MID)
        cell.border = B_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[col_letter].width = width

    samples = [
        ("Darn Tough Wool Socks (Sample)", "Apparel", "Owned", 20, 0, 0, 50, "Year", 4, 5, "Lifetime guarantee; $0.10/use benchmark"),
        ("Ergonomic Office Chair (Sample)", "Home & Furniture", "Owned", 900, 0, 250, 5, "Week", 10, 5, "High daily use; strong resale value"),
        ("Cast Iron Skillet (Sample)", "Tools", "Owned", 45, 0, 10, 4, "Week", 25, 5, "Indestructible; improves with patina"),
        ("Streaming Bundle (Sample)", "Subscription / Service", "Owned", 0, 180, 0, 2, "Week", 1, 3, "Pure subscription ($15/mo = $180/yr)"),
        ("Noise-Canceling Headphones (Sample)", "Tech & Gear", "Considering", 350, 15, 50, 4, "Week", 4, 4, "Includes $15/yr ear-pad replacement"),
        ("Espresso Machine (Sample)", "Hobby", "Considering", 700, 60, 200, 1, "Day", 7, 4, "Replaceable gaskets; $60/yr descaler & filter"),
    ]

    for r in range(9, max_row + 1):
        ws.row_dimensions[r].height = 24
        is_even = (r % 2 == 0)
        sample_idx = r - 9
        sample = samples[sample_idx] if sample_idx < len(samples) else None

        input_map = {}
        if sample:
            input_map = {
                "B": sample[0], "C": sample[1], "D": sample[2],
                "E": sample[3], "F": sample[4], "G": sample[5],
                "H": sample[6], "I": sample[7], "J": sample[8],
                "K": sample[9], "S": sample[10],
            }

        for col_letter, _, _, is_input, num_fmt, align_h in columns:
            cell = ws[f"{col_letter}{r}"]
            cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0)
            if num_fmt:
                cell.number_format = num_fmt

            if is_input:
                cell.fill = FILL_INPUT_ALT if is_even else FILL_INPUT
                cell.font = F_INPUT_BOLD if col_letter == "B" else F_INPUT
                cell.border = B_INPUT
                if col_letter in input_map:
                    cell.value = input_map[col_letter]
            else:
                if col_letter == "O":
                    cell.fill = FILL_EMERALD
                    cell.font = F_OUT_HERO
                else:
                    cell.fill = FILL_ZEBRA if is_even else FILL_WHITE
                    cell.font = F_OUT
                cell.border = B_CELL

        # Formulas L..R
        ws[f"L{r}"] = f'=IF(OR(B{r}="",H{r}="",I{r}=""),"",IF(I{r}="Day",H{r}*365,IF(I{r}="Week",H{r}*52,IF(I{r}="Month",H{r}*12,IF(I{r}="Year",H{r},"")))))'
        ws[f"M{r}"] = f'=IF(OR(L{r}="",J{r}=""),"",L{r}*J{r})'
        ws[f"N{r}"] = f'=IF(OR(B{r}="",E{r}=""),"",E{r}+IF(F{r}="",0,F{r})*IF(J{r}="",0,J{r})-IF(G{r}="",0,G{r}))'
        ws[f"O{r}"] = f'=IF(OR(N{r}="",M{r}="",M{r}=0),"",N{r}/M{r})'
        ws[f"P{r}"] = f'=IF(OR(N{r}="",J{r}="",J{r}=0),"",N{r}/J{r})'
        ws[f"Q{r}"] = f'=IF(P{r}="","",P{r}/12)'
        ws[f"R{r}"] = f'=IF(O{r}="","",IF(O{r}<=0.25,"Exceptional",IF(O{r}<=1,"Great Value",IF(O{r}<=5,"Moderate","High $/Use"))))'

    # Data validations
    dv_cat = DataValidation(type="list", formula1=f'"{",".join(CATEGORIES)}"', allow_blank=True)
    dv_stat = DataValidation(type="list", formula1=f'"{",".join(STATUSES)}"', allow_blank=True)
    dv_unit = DataValidation(type="list", formula1=f'"{",".join(FREQ_UNITS)}"', allow_blank=True)
    dv_rate = DataValidation(type="list", formula1=f'"{",".join(RATINGS)}"', allow_blank=True)
    for dv, rng in [
        (dv_cat, f"C9:C{max_row}"),
        (dv_stat, f"D9:D{max_row}"),
        (dv_unit, f"I9:I{max_row}"),
        (dv_rate, f"K9:K{max_row}"),
    ]:
        ws.add_data_validation(dv)
        dv.add(rng)

    # Conditional formatting: Status pills (D9:D108)
    status_styles = [
        ("Owned", "DCFCE7", "166534"),
        ("Considering", "FEF3C7", "92400E"),
        ("Retired", "E2E8F0", "475569"),
        ("Skipped", "FEE2E2", "991B1B"),
    ]
    for val, bg, fg in status_styles:
        ws.conditional_formatting.add(
            f"D9:D{max_row}",
            CellIsRule(
                operator="equal",
                formula=[f'"{val}"'],
                fill=PatternFill("solid", fgColor=bg),
                font=Font(name=FONT_FAMILY, size=9.5, bold=True, color=fg),
            ),
        )

    # Conditional formatting: Durability Rating pills (K9:K108)
    rating_styles = [
        ("5", "DCFCE7", "166534"),
        ("4", "ECFDF5", "047857"),
        ("3", "FEF3C7", "92400E"),
        ("2", "FFEDD5", "9A3412"),
        ("1", "FEE2E2", "991B1B"),
    ]
    for val, bg, fg in rating_styles:
        ws.conditional_formatting.add(
            f"K9:K{max_row}",
            CellIsRule(
                operator="equal",
                formula=[val],
                fill=PatternFill("solid", fgColor=bg),
                font=Font(name=FONT_FAMILY, size=9.5, bold=True, color=fg),
            ),
        )

    # Conditional formatting: Value Tier pills (R9:R108)
    tier_styles = [
        ("Exceptional", "D1FAE5", "065F46"),
        ("Great Value", "E0F2FE", "075985"),
        ("Moderate", "FEF3C7", "92400E"),
        ("High $/Use", "FEE2E2", "991B1B"),
    ]
    for val, bg, fg in tier_styles:
        ws.conditional_formatting.add(
            f"R9:R{max_row}",
            CellIsRule(
                operator="equal",
                formula=[f'"{val}"'],
                fill=PatternFill("solid", fgColor=bg),
                font=Font(name=FONT_FAMILY, size=9.5, bold=True, color=fg),
            ),
        )

    # Conditional formatting: Green -> Amber -> Red scale on Cost / Use (O9:O108)
    ws.conditional_formatting.add(
        f"O9:O{max_row}",
        ColorScaleRule(
            start_type="min", start_color="D1FAE5",
            mid_type="percentile", mid_value=50, mid_color="FEF3C7",
            end_type="max", end_color="FEE2E2",
        ),
    )

    ws.auto_filter.ref = f"B8:S{max_row}"


def build_insights_tab(wb: Workbook):
    ws = wb.create_sheet(title="Insights & Charts")
    ws.views.sheetView[0].showGridLines = False
    ws.freeze_panes = None

    fill_canvas(ws, max_row=48, max_col=15)

    ws.row_dimensions[1].height = 12
    ws.column_dimensions["A"].width = 2.5
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 2.5
    for col in ["I", "J", "K", "L", "M", "N"]:
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["O"].width = 2.5

    # Header Banner (B2:N3)
    ws.merge_cells("B2:N2")
    ws["B2"] = "Portfolio Insights & Visualizations"
    style_range(
        ws, "B2:N2", font=F_TITLE, fill=FILL_DARK,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[2].height = 36

    ws.merge_cells("B3:N3")
    ws["B3"] = "Live rollups and charts synced automatically with your Master Tracker."
    style_range(
        ws, "B3:N3", font=F_SUBTITLE, fill=FILL_MID,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 12

    mt = "'Master Tracker'"

    # =========================================================================
    # Table 1: Category Breakdown (B5:G15)
    # =========================================================================
    ws.merge_cells("B5:G5")
    ws["B5"] = "1. BREAKDOWN BY CATEGORY"
    style_range(
        ws, "B5:G5", font=F_SECTION, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[5].height = 24

    cat_headers = [
        ("B6", "Category", "left"),
        ("C6", "Items", "center"),
        ("D6", "Total Net Cost", "right"),
        ("E6", "Annualized Cost", "right"),
        ("F6", "Avg $/Use", "right"),
        ("G6", "Avg Rating", "center"),
    ]
    ws.row_dimensions[6].height = 28
    for coord, text, align_h in cat_headers:
        cell = ws[coord]
        cell.value = text
        cell.font = F_TH
        cell.fill = FILL_MID
        cell.border = B_HEADER
        cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0, wrap_text=True)

    for idx, cat in enumerate(CATEGORIES):
        r = 7 + idx
        ws.row_dimensions[r].height = 23
        r_fill = FILL_ZEBRA if (r % 2 == 0) else FILL_WHITE

        ws[f"B{r}"] = cat
        ws[f"C{r}"] = f'=COUNTIF({mt}!$C$9:$C$108,B{r})'
        ws[f"D{r}"] = f'=SUMIF({mt}!$C$9:$C$108,B{r},{mt}!$N$9:$N$108)'
        ws[f"E{r}"] = f'=SUMIF({mt}!$C$9:$C$108,B{r},{mt}!$P$9:$P$108)'
        ws[f"F{r}"] = f'=IF(C{r}=0,"",AVERAGEIF({mt}!$C$9:$C$108,B{r},{mt}!$O$9:$O$108))'
        ws[f"G{r}"] = f'=IF(C{r}=0,"",AVERAGEIF({mt}!$C$9:$C$108,B{r},{mt}!$K$9:$K$108))'

        styles = [
            ("B", None, "left"),
            ("C", FMT_INT, "center"),
            ("D", FMT_CURRENCY, "right"),
            ("E", FMT_CURRENCY, "right"),
            ("F", FMT_CURRENCY, "right"),
            ("G", FMT_DEC1, "center"),
        ]
        for col_letter, fmt, align_h in styles:
            cell = ws[f"{col_letter}{r}"]
            cell.font = F_LABEL if col_letter == "B" else F_OUT
            cell.fill = r_fill
            cell.border = B_CELL
            cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0)
            if fmt:
                cell.number_format = fmt

    # Total row (Row 15)
    ws.row_dimensions[15].height = 25
    ws["B15"] = "Total / Overall"
    ws["C15"] = "=SUM(C7:C14)"
    ws["D15"] = "=SUM(D7:D14)"
    ws["E15"] = "=SUM(E7:E14)"
    ws["F15"] = f'=IFERROR(AVERAGE({mt}!$O$9:$O$108),"")'
    ws["G15"] = f'=IFERROR(AVERAGE({mt}!$K$9:$K$108),"")'
    for col_letter, fmt, align_h in [
        ("B", None, "left"), ("C", FMT_INT, "center"), ("D", FMT_CURRENCY, "right"),
        ("E", FMT_CURRENCY, "right"), ("F", FMT_CURRENCY, "right"), ("G", FMT_DEC1, "center"),
    ]:
        cell = ws[f"{col_letter}15"]
        cell.font = F_LABEL_BOLD
        cell.fill = FILL_EMERALD
        cell.border = B_HIGHLIGHT
        cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0)
        if fmt:
            cell.number_format = fmt

    ws.row_dimensions[16].height = 12

    # =========================================================================
    # Table 2: Ownership Status Summary (B17:G22)
    # =========================================================================
    ws.merge_cells("B17:G17")
    ws["B17"] = "2. BREAKDOWN BY OWNERSHIP STATUS"
    style_range(
        ws, "B17:G17", font=F_SECTION, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[17].height = 24

    ws.row_dimensions[18].height = 26
    for coord, text, align_h in [
        ("B18", "Status", "left"), ("C18", "Items", "center"),
        ("D18", "Total Net Cost", "right"), ("E18", "Annualized Cost", "right"),
        ("F18", "Avg $/Use", "right"), ("G18", "% of Items", "center"),
    ]:
        cell = ws[coord]
        cell.value = text
        cell.font = F_TH
        cell.fill = FILL_MID
        cell.border = B_HEADER
        cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0)

    for idx, status in enumerate(STATUSES):
        r = 19 + idx
        ws.row_dimensions[r].height = 23
        r_fill = FILL_ZEBRA if (r % 2 == 0) else FILL_WHITE

        ws[f"B{r}"] = status
        ws[f"C{r}"] = f'=COUNTIF({mt}!$D$9:$D$108,B{r})'
        ws[f"D{r}"] = f'=SUMIF({mt}!$D$9:$D$108,B{r},{mt}!$N$9:$N$108)'
        ws[f"E{r}"] = f'=SUMIF({mt}!$D$9:$D$108,B{r},{mt}!$P$9:$P$108)'
        ws[f"F{r}"] = f'=IF(C{r}=0,"",AVERAGEIF({mt}!$D$9:$D$108,B{r},{mt}!$O$9:$O$108))'
        ws[f"G{r}"] = f'=IF($C$15=0,"",C{r}/$C$15)'

        for col_letter, fmt, align_h in [
            ("B", None, "left"), ("C", FMT_INT, "center"), ("D", FMT_CURRENCY, "right"),
            ("E", FMT_CURRENCY, "right"), ("F", FMT_CURRENCY, "right"), ("G", FMT_PCT, "center"),
        ]:
            cell = ws[f"{col_letter}{r}"]
            cell.font = F_LABEL_BOLD if col_letter == "B" else F_OUT
            cell.fill = r_fill
            cell.border = B_CELL
            cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0)
            if fmt:
                cell.number_format = fmt

    ws.row_dimensions[23].height = 12

    # =========================================================================
    # Table 3: Cost-per-Use Value Tiers (B24:G29)
    # =========================================================================
    ws.merge_cells("B24:G24")
    ws["B24"] = "3. COST-PER-USE VALUE TIERS"
    style_range(
        ws, "B24:G24", font=F_SECTION, fill=FILL_SECTION, border=B_CELL,
        alignment=Alignment(horizontal="left", vertical="center", indent=1),
    )
    ws.row_dimensions[24].height = 24

    ws.row_dimensions[25].height = 26
    ws.merge_cells("C25:D25")
    for coord, text, align_h in [
        ("B25", "Value Tier", "left"), ("C25", "Threshold", "center"),
        ("E25", "Items", "center"), ("F25", "% of Tracked", "center"), ("G25", "Total Spend", "right"),
    ]:
        cell = ws[coord]
        cell.value = text
        cell.font = F_TH
        cell.fill = FILL_MID
        cell.border = B_HEADER
        cell.alignment = Alignment(horizontal=align_h, vertical="center", indent=1 if align_h == "left" else 0)
    ws["D25"].fill = FILL_MID
    ws["D25"].border = B_HEADER

    tiers = [
        ("Exceptional", "<= $0.25 / use", "D1FAE5", "065F46"),
        ("Great Value", "$0.26 - $1.00 / use", "E0F2FE", "075985"),
        ("Moderate", "$1.01 - $5.00 / use", "FEF3C7", "92400E"),
        ("High $/Use", "> $5.00 / use", "FEE2E2", "991B1B"),
    ]
    for idx, (tier_name, thresh, pill_bg, pill_fg) in enumerate(tiers):
        r = 26 + idx
        ws.row_dimensions[r].height = 24
        ws.merge_cells(f"C{r}:D{r}")

        ws[f"B{r}"] = tier_name
        ws[f"B{r}"].font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=pill_fg)
        ws[f"B{r}"].fill = PatternFill("solid", fgColor=pill_bg)
        ws[f"B{r}"].border = B_CELL
        ws[f"B{r}"].alignment = Alignment(horizontal="center", vertical="center")

        ws[f"C{r}"] = thresh
        style_range(
            ws, f"C{r}:D{r}", font=F_NOTE, fill=FILL_WHITE, border=B_CELL,
            alignment=Alignment(horizontal="center", vertical="center"),
        )

        ws[f"E{r}"] = f'=COUNTIF({mt}!$R$9:$R$108,B{r})'
        ws[f"E{r}"].font = F_OUT
        ws[f"E{r}"].fill = FILL_WHITE
        ws[f"E{r}"].border = B_CELL
        ws[f"E{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"E{r}"].number_format = FMT_INT

        ws[f"F{r}"] = f'=IF($C$15=0,"",E{r}/$C$15)'
        ws[f"F{r}"].font = F_OUT
        ws[f"F{r}"].fill = FILL_WHITE
        ws[f"F{r}"].border = B_CELL
        ws[f"F{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"F{r}"].number_format = FMT_PCT

        ws[f"G{r}"] = f'=SUMIF({mt}!$R$9:$R$108,B{r},{mt}!$N$9:$N$108)'
        ws[f"G{r}"].font = F_OUT
        ws[f"G{r}"].fill = FILL_WHITE
        ws[f"G{r}"].border = B_CELL
        ws[f"G{r}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"G{r}"].number_format = FMT_CURRENCY

    ws.row_dimensions[30].height = 14

    # =========================================================================
    # 4 NATIVE PORTFOLIO CHARTS
    # =========================================================================
    # Chart 1 (Top-Right, I5): Horizontal Bar — Total Net Cost by Category
    c1 = BarChart()
    c1.type = "bar"
    c1.grouping = "clustered"
    style_chart_common(c1, "Total Net Cost by Category ($)", y_title=None, y_fmt="$#,##0")
    c1.legend = None
    c1.dataLabels = DataLabelList()
    c1.dataLabels.showVal = True
    c1.width = 16.4
    c1.height = 9.4

    c1.add_data(Reference(ws, min_col=4, min_row=6, max_row=14), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=2, min_row=7, max_row=14))
    c1.series[0].graphicalProperties.solidFill = C_OPT_B
    c1.series[0].graphicalProperties.line.noFill = True
    ws.add_chart(c1, "I5")

    # Chart 2 (Mid-Right, I18): Horizontal Bar — Avg Cost per Use by Category
    c2 = BarChart()
    c2.type = "bar"
    c2.grouping = "clustered"
    style_chart_common(c2, "Average Cost per Use by Category ($)", y_title=None, y_fmt="$#,##0.00")
    c2.legend = None
    c2.dataLabels = DataLabelList()
    c2.dataLabels.showVal = True
    c2.width = 16.4
    c2.height = 9.4

    c2.add_data(Reference(ws, min_col=6, min_row=6, max_row=14), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=2, min_row=7, max_row=14))
    c2.series[0].graphicalProperties.solidFill = C_OPT_C
    c2.series[0].graphicalProperties.line.noFill = True
    ws.add_chart(c2, "I18")

    # Chart 3 (Bottom-Left, B31): Column Chart — Items by Value Tier
    c3 = BarChart()
    c3.type = "col"
    c3.grouping = "clustered"
    style_chart_common(c3, "Items by Cost-per-Use Value Tier", y_title="Items", y_fmt="#,##0")
    c3.legend = None
    c3.dataLabels = DataLabelList()
    c3.dataLabels.showVal = True
    c3.width = 18.8
    c3.height = 8.4

    c3.add_data(Reference(ws, min_col=5, min_row=25, max_row=29), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=2, min_row=26, max_row=29))
    c3.series[0].graphicalProperties.solidFill = C_OPT_C
    c3.series[0].graphicalProperties.line.noFill = True
    for i, hex_color in enumerate(["059669", "0284C7", "D97706", "DC2626"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = hex_color
        pt.graphicalProperties.line.noFill = True
        c3.series[0].data_points.append(pt)
    ws.add_chart(c3, "B31")

    # Chart 4 (Bottom-Right, I31): Column Chart — Total Spend by Ownership Status
    c4 = BarChart()
    c4.type = "col"
    c4.grouping = "clustered"
    style_chart_common(c4, "Total Net Cost by Ownership Status ($)", y_title="Net Spend ($)", y_fmt="$#,##0")
    c4.legend = None
    c4.dataLabels = DataLabelList()
    c4.dataLabels.showVal = True
    c4.width = 16.4
    c4.height = 8.4

    c4.add_data(Reference(ws, min_col=4, min_row=18, max_row=22), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=2, min_row=19, max_row=22))
    c4.series[0].graphicalProperties.solidFill = C_OPT_A
    c4.series[0].graphicalProperties.line.noFill = True
    for i, hex_color in enumerate([C_RENTAL, C_OPT_C, "64748B", "EF4444"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = hex_color
        pt.graphicalProperties.line.noFill = True
        c4.series[0].data_points.append(pt)
    ws.add_chart(c4, "I31")


def main():
    out_dir = Path(__file__).resolve().parent
    out_path = out_dir / "Cost_Per_Use_Calculator.xlsx"

    wb = Workbook()
    build_quick_calc_tab(wb)
    build_calculator_tab(wb)
    build_master_tracker_tab(wb)
    build_insights_tab(wb)
    wb.save(out_path)
    print(f"Saved upgraded workbook to: {out_path}")


if __name__ == "__main__":
    main()
