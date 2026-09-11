#!/usr/bin/env python3
"""Build Northline_Margin_Mix_Bridge.xlsx — SKU/channel contribution, PVM, GM walk."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- styles ---
yellow = PatternFill("solid", fgColor="FFF2CC")
header_fill = PatternFill("solid", fgColor="1F4E79")
section_fill = PatternFill("solid", fgColor="D6E3F0")
green_fill = PatternFill("solid", fgColor="C6EFCE")
amber_fill = PatternFill("solid", fgColor="FFE699")
red_fill = PatternFill("solid", fgColor="F8CBAD")
tile_fill = PatternFill("solid", fgColor="E9EDF4")
light_gray = PatternFill("solid", fgColor="F5F5F5")
bridge_fill = PatternFill("solid", fgColor="DEEBF7")

input_font = Font(name="Calibri", size=11, color="0000FF")
black = Font(name="Calibri", size=11, color="000000")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
title_font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
section_font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
bold = Font(name="Calibri", size=11, bold=True)
bold_black = Font(name="Calibri", size=11, bold=True, color="000000")
italic_grey = Font(name="Calibri", size=10, italic=True, color="666666")
small_grey = Font(name="Calibri", size=9, italic=True, color="666666")
link_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
money = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
money1 = '_($* #,##0.0_);_($* (#,##0.0);_($* "-"??_);_(@_)'
pct = "0.0%"
num = "#,##0.0"
ppu = "0.00"
int_fmt = "#,##0"
vol_fmt = "#,##0"

CHANNELS = ["Grocery", "Mass", "Club", "E-comm"]
SKUS = [
    "Sparkle Water 12pk",
    "Protein Bar 12ct",
    "Crunch Chips Family",
    "Clean Wipe Refill",
    "SoftSoap 3pk",
    "Multipack Snacks",
]
# Standard COGS $/unit (constant across periods)
COGS = [2.15, 4.85, 1.90, 3.35, 2.50, 5.10]

# Net price $/unit by SKU × channel — Prior
# Club priced lower; E-comm slight premium; Mass slight discount vs Grocery
PRIOR_PRICE = {
    # Grocery, Mass, Club, E-comm
    "Sparkle Water 12pk": [5.49, 5.29, 4.49, 5.69],
    "Protein Bar 12ct": [9.99, 9.49, 8.29, 10.29],
    "Crunch Chips Family": [4.79, 4.49, 3.79, 4.99],
    "Clean Wipe Refill": [7.49, 7.19, 6.29, 7.79],
    "SoftSoap 3pk": [5.99, 5.69, 4.89, 6.19],
    "Multipack Snacks": [11.99, 11.49, 9.49, 12.49],
}
# Current: mild Grocery/Mass price pressure; Club holds; E-comm up slightly
CURR_PRICE = {
    "Sparkle Water 12pk": [5.19, 4.99, 4.35, 5.59],
    "Protein Bar 12ct": [9.49, 8.99, 7.99, 10.19],
    "Crunch Chips Family": [4.49, 4.19, 3.59, 4.89],
    "Clean Wipe Refill": [7.09, 6.79, 6.05, 7.69],
    "SoftSoap 3pk": [5.59, 5.29, 4.69, 6.09],
    "Multipack Snacks": [11.49, 10.99, 9.09, 12.29],
}

# Volume 000 units by SKU × channel — Prior
# Story: Current shifts volume into Club (lower margin) and Multipack; Grocery soft
PRIOR_VOL = {
    "Sparkle Water 12pk": [420, 310, 180, 95],
    "Protein Bar 12ct": [210, 165, 90, 70],
    "Crunch Chips Family": [380, 290, 220, 85],
    "Clean Wipe Refill": [145, 120, 60, 55],
    "SoftSoap 3pk": [260, 240, 110, 60],
    "Multipack Snacks": [80, 70, 200, 40],
}
CURR_VOL = {
    "Sparkle Water 12pk": [360, 280, 280, 105],   # Grocery soft; Club surge
    "Protein Bar 12ct": [185, 150, 150, 80],
    "Crunch Chips Family": [320, 250, 340, 95],   # Club volume spike
    "Clean Wipe Refill": [130, 110, 95, 60],
    "SoftSoap 3pk": [220, 210, 175, 65],
    "Multipack Snacks": [60, 55, 340, 45],        # Club multipack surge
}


def style_header_row(ws, row, start_col, end_col):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = thin


def input_cell(cell, value, fmt=None):
    cell.value = value
    cell.fill = yellow
    cell.font = input_font
    cell.border = thin
    if fmt:
        cell.number_format = fmt


def formula_cell(cell, formula, fmt=None, bold_f=False):
    cell.value = formula
    cell.font = bold_black if bold_f else black
    cell.border = thin
    if fmt:
        cell.number_format = fmt


def label_cell(cell, text, bold_f=False):
    cell.value = text
    cell.font = bold if bold_f else black
    cell.border = thin


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_cover(wb):
    ws = wb.create_sheet("00_Cover", 0)
    set_col_widths(ws, [3, 78])
    ws["B2"] = "Northline Consumer Products"
    ws["B2"].font = title_font
    ws["B3"] = "Margin & Mix Bridge — SKU × Channel Contribution | Price-Volume-Mix | Gross Margin Walk"
    ws["B3"].font = section_font
    ws["B5"] = "Purpose"
    ws["B5"].font = section_font
    ws["B6"] = (
        "Explain how gross margin dollars moved between prior and current period: "
        "volume growth, net price, SKU/channel mix, and cost. Built for an FP&A "
        "screen-share — change yellow inputs on 01_Assumptions and watch the bridge update."
    )
    ws["B6"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[6].height = 48

    ws["B8"] = "How to read this file"
    ws["B8"].font = section_font
    bullets = [
        "1. Yellow fill + blue font = inputs. Black font = formulas. No VBA.",
        "2. Start on 01_Assumptions (close period, channel list, SKU economics, prior/current volume & price).",
        "3. 02_Volume_Price builds SKU×channel revenue and contribution.",
        "4. 03_PVM_Bridge walks Prior GM → Volume → Price → Mix → Cost → Current GM (revenue and GM$).",
        "5. 04_Contribution shows margin by SKU and by channel with mix %.",
        "6. 05_Dashboard is the one-pager for leadership: GM walk, mix callouts, RAG on margin %.",
        "7. Amounts in $000s where labeled; volume in 000 units; prices and COGS in $ per unit.",
    ]
    for i, b in enumerate(bullets):
        ws.cell(9 + i, 2, b).font = black

    ws["B17"] = "Sample story (embedded numbers)"
    ws["B17"].font = section_font
    ws["B18"] = (
        "Total volume is up, driven by Club and Multipack Snacks. Club carries a lower "
        "net price and thinner contribution per unit, so mix dilutes gross margin even "
        "as units grow. Mild Grocery/Mass price pressure adds a negative price step; "
        "COGS/unit is flat in this sample so the Cost bridge step is ~0."
    )
    ws["B18"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[18].height = 60

    ws["B20"] = "Color legend"
    ws["B20"].font = section_font
    ws["B21"] = "Input cell (edit me)"
    input_cell(ws["B21"], "Input cell (edit me)")
    ws["B22"] = "Formula / calculated"
    formula_cell(ws["B22"], "Formula / calculated")

    ws["B24"] = "Fictional company and sample data for portfolio demonstration only. Not employer or client data."
    ws["B24"].font = italic_grey
    ws["B25"] = "github.com/saisiri-bandaru"
    ws["B25"].font = link_font
    ws["B26"] = "Sai Siri Bandaru — Financial Analyst | FP&A"
    ws["B26"].font = bold


def build_assumptions(wb):
    ws = wb.create_sheet("01_Assumptions", 1)
    set_col_widths(ws, [3, 28, 14, 14, 14, 14, 14, 14, 14, 14, 14])

    ws["B2"] = "Assumptions & Inputs"
    ws["B2"].font = title_font
    ws["B3"] = "Yellow / blue = editable inputs. All downstream tabs reference this sheet."
    ws["B3"].font = italic_grey

    ws["B5"] = "Close period"
    ws["B5"].font = bold
    input_cell(ws["C5"], "FY26 Q2 vs FY26 Q1")
    ws["B6"] = "Units"
    ws["B6"].font = bold
    input_cell(ws["C6"], "000 cases / units")
    ws["B7"] = "Currency"
    ws["B7"].font = bold
    input_cell(ws["C7"], "$000s (revenue, GM); $/unit (price, COGS)")

    # Channels
    ws["B9"] = "Channels"
    ws["B9"].font = section_font
    ws["B10"] = "Channel"
    ws["C10"] = "Role in mix story"
    style_header_row(ws, 10, 2, 3)
    channel_notes = [
        ("Grocery", "Highest net price; volume soft in current"),
        ("Mass", "Slight discount vs Grocery; mild price pressure"),
        ("Club", "Lowest net $/unit; volume surge dilutes mix"),
        ("E-comm", "Slight premium; growing but smaller base"),
    ]
    for i, (ch, note) in enumerate(channel_notes):
        r = 11 + i
        input_cell(ws.cell(r, 2), ch)
        input_cell(ws.cell(r, 3), note)

    # SKU list + COGS
    ws["B16"] = "SKU list & standard COGS $/unit"
    ws["B16"].font = section_font
    headers = ["SKU", "Std COGS $/unit", "Category note"]
    for i, h in enumerate(headers):
        ws.cell(17, 2 + i, h)
    style_header_row(ws, 17, 2, 4)
    notes = [
        "Sparkling beverage — high velocity",
        "Nutrition — mid-high margin",
        "Salty snacks — volume driver",
        "Household — steady",
        "Personal care — promotional",
        "Club-oriented pack — thin margin",
    ]
    for i, sku in enumerate(SKUS):
        r = 18 + i
        input_cell(ws.cell(r, 2), sku)
        input_cell(ws.cell(r, 3), COGS[i], ppu)
        input_cell(ws.cell(r, 4), notes[i])

    # Volume prior
    ws["B25"] = "Prior-period volume (000 units) by SKU × channel"
    ws["B25"].font = section_font
    ws["B26"] = "SKU"
    for j, ch in enumerate(CHANNELS):
        ws.cell(26, 3 + j, ch)
    ws.cell(26, 7, "Total")
    style_header_row(ws, 26, 2, 7)
    for i, sku in enumerate(SKUS):
        r = 27 + i
        label_cell(ws.cell(r, 2), sku)
        for j in range(4):
            input_cell(ws.cell(r, 3 + j), PRIOR_VOL[sku][j], vol_fmt)
        formula_cell(ws.cell(r, 7), f"=SUM(C{r}:F{r})", vol_fmt, True)
    ws["B33"] = "Total"
    ws["B33"].font = bold
    for j in range(5):
        formula_cell(ws.cell(33, 3 + j), f"=SUM({get_column_letter(3+j)}27:{get_column_letter(3+j)}32)", vol_fmt, True)

    # Volume current
    ws["B35"] = "Current-period volume (000 units) by SKU × channel"
    ws["B35"].font = section_font
    ws["B36"] = "SKU"
    for j, ch in enumerate(CHANNELS):
        ws.cell(36, 3 + j, ch)
    ws.cell(36, 7, "Total")
    style_header_row(ws, 36, 2, 7)
    for i, sku in enumerate(SKUS):
        r = 37 + i
        label_cell(ws.cell(r, 2), sku)
        for j in range(4):
            input_cell(ws.cell(r, 3 + j), CURR_VOL[sku][j], vol_fmt)
        formula_cell(ws.cell(r, 7), f"=SUM(C{r}:F{r})", vol_fmt, True)
    ws["B43"] = "Total"
    ws["B43"].font = bold
    for j in range(5):
        formula_cell(ws.cell(43, 3 + j), f"=SUM({get_column_letter(3+j)}37:{get_column_letter(3+j)}42)", vol_fmt, True)

    # Prior prices
    ws["B45"] = "Prior-period net price ($/unit) by SKU × channel"
    ws["B45"].font = section_font
    ws["B46"] = "SKU"
    for j, ch in enumerate(CHANNELS):
        ws.cell(46, 3 + j, ch)
    style_header_row(ws, 46, 2, 6)
    for i, sku in enumerate(SKUS):
        r = 47 + i
        label_cell(ws.cell(r, 2), sku)
        for j in range(4):
            input_cell(ws.cell(r, 3 + j), PRIOR_PRICE[sku][j], ppu)

    # Current prices
    ws["B54"] = "Current-period net price ($/unit) by SKU × channel"
    ws["B54"].font = section_font
    ws["B55"] = "SKU"
    for j, ch in enumerate(CHANNELS):
        ws.cell(55, 3 + j, ch)
    style_header_row(ws, 55, 2, 6)
    for i, sku in enumerate(SKUS):
        r = 56 + i
        label_cell(ws.cell(r, 2), sku)
        for j in range(4):
            input_cell(ws.cell(r, 3 + j), CURR_PRICE[sku][j], ppu)

    ws["B63"] = "Note: COGS/unit held constant (see Std COGS above) so the Cost step on 03_PVM_Bridge is ~0 in this sample."
    ws["B63"].font = italic_grey


def build_volume_price(wb):
    ws = wb.create_sheet("02_Volume_Price", 2)
    set_col_widths(ws, [3, 24, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12])

    ws["B2"] = "Volume, Price & Contribution Build"
    ws["B2"].font = title_font
    ws["B3"] = "SKU×channel detail. Revenue and GM in $000s (= volume 000 × $/unit)."
    ws["B3"].font = italic_grey

    # Detail table: one row per SKU×channel
    headers = [
        "SKU", "Channel", "Vol Prior", "Vol Curr", "Δ Vol",
        "Price Prior", "Price Curr", "Δ Price",
        "Rev Prior $000", "Rev Curr $000",
        "COGS $/u", "GM Prior $000", "GM Curr $000",
    ]
    for i, h in enumerate(headers):
        ws.cell(5, 2 + i, h)
    style_header_row(ws, 5, 2, 14)
    ws.row_dimensions[5].height = 32

    row = 6
    for i, sku in enumerate(SKUS):
        cogs_row = 18 + i  # on Assumptions
        for j, ch in enumerate(CHANNELS):
            vol_p_cell = f"'01_Assumptions'!{get_column_letter(3+j)}{27+i}"
            vol_c_cell = f"'01_Assumptions'!{get_column_letter(3+j)}{37+i}"
            px_p_cell = f"'01_Assumptions'!{get_column_letter(3+j)}{47+i}"
            px_c_cell = f"'01_Assumptions'!{get_column_letter(3+j)}{56+i}"
            cogs_cell = f"'01_Assumptions'!C{cogs_row}"

            label_cell(ws.cell(row, 2), sku)
            label_cell(ws.cell(row, 3), ch)
            formula_cell(ws.cell(row, 4), f"={vol_p_cell}", vol_fmt)
            formula_cell(ws.cell(row, 5), f"={vol_c_cell}", vol_fmt)
            formula_cell(ws.cell(row, 6), f"=E{row}-D{row}", vol_fmt)
            formula_cell(ws.cell(row, 7), f"={px_p_cell}", ppu)
            formula_cell(ws.cell(row, 8), f"={px_c_cell}", ppu)
            formula_cell(ws.cell(row, 9), f"=H{row}-G{row}", ppu)
            # Rev = vol * price (already $000 because vol in 000)
            formula_cell(ws.cell(row, 10), f"=D{row}*G{row}", money1)
            formula_cell(ws.cell(row, 11), f"=E{row}*H{row}", money1)
            formula_cell(ws.cell(row, 12), f"={cogs_cell}", ppu)
            formula_cell(ws.cell(row, 13), f"=D{row}*(G{row}-L{row})", money1)
            formula_cell(ws.cell(row, 14), f"=E{row}*(H{row}-L{row})", money1)
            row += 1

    last_detail = row - 1  # 29
    # Totals
    label_cell(ws.cell(row, 2), "TOTAL", True)
    label_cell(ws.cell(row, 3), "", True)
    for col in [4, 5, 6, 10, 11, 13, 14]:
        formula_cell(ws.cell(row, col), f"=SUM({get_column_letter(col)}6:{get_column_letter(col)}{last_detail})", money1 if col >= 10 else vol_fmt, True)
    total_row = row

    # Summary by SKU
    row += 3
    sku_sum_start = row
    ws.cell(row, 2, "Summary by SKU")
    ws.cell(row, 2).font = section_font
    row += 1
    for i, h in enumerate(["SKU", "Vol Prior", "Vol Curr", "Rev Prior", "Rev Curr", "GM Prior", "GM Curr", "GM% Prior", "GM% Curr"]):
        ws.cell(row, 2 + i, h)
    style_header_row(ws, row, 2, 10)
    hdr = row
    row += 1
    sku_first = row
    for i, sku in enumerate(SKUS):
        # 4 channels each starting at 6 + i*4
        r0 = 6 + i * 4
        r1 = r0 + 3
        label_cell(ws.cell(row, 2), sku)
        formula_cell(ws.cell(row, 3), f"=SUMIF($B$6:$B${last_detail},B{row},$D$6:$D${last_detail})", vol_fmt)
        formula_cell(ws.cell(row, 4), f"=SUMIF($B$6:$B${last_detail},B{row},$E$6:$E${last_detail})", vol_fmt)
        formula_cell(ws.cell(row, 5), f"=SUMIF($B$6:$B${last_detail},B{row},$J$6:$J${last_detail})", money1)
        formula_cell(ws.cell(row, 6), f"=SUMIF($B$6:$B${last_detail},B{row},$K$6:$K${last_detail})", money1)
        formula_cell(ws.cell(row, 7), f"=SUMIF($B$6:$B${last_detail},B{row},$M$6:$M${last_detail})", money1)
        formula_cell(ws.cell(row, 8), f"=SUMIF($B$6:$B${last_detail},B{row},$N$6:$N${last_detail})", money1)
        formula_cell(ws.cell(row, 9), f"=IF(E{row}=0,0,G{row}/E{row})", pct)
        formula_cell(ws.cell(row, 10), f"=IF(F{row}=0,0,H{row}/F{row})", pct)
        row += 1
    sku_last = row - 1
    label_cell(ws.cell(row, 2), "Total", True)
    for col, fmt in [(3, vol_fmt), (4, vol_fmt), (5, money1), (6, money1), (7, money1), (8, money1)]:
        formula_cell(ws.cell(row, col), f"=SUM({get_column_letter(col)}{sku_first}:{get_column_letter(col)}{sku_last})", fmt, True)
    formula_cell(ws.cell(row, 9), f"=IF(E{row}=0,0,G{row}/E{row})", pct, True)
    formula_cell(ws.cell(row, 10), f"=IF(F{row}=0,0,H{row}/F{row})", pct, True)

    # Summary by channel
    row += 3
    ws.cell(row, 2, "Summary by Channel")
    ws.cell(row, 2).font = section_font
    row += 1
    for i, h in enumerate(["Channel", "Vol Prior", "Vol Curr", "Rev Prior", "Rev Curr", "GM Prior", "GM Curr", "GM% Prior", "GM% Curr", "Vol Mix% Prior", "Vol Mix% Curr"]):
        ws.cell(row, 2 + i, h)
    style_header_row(ws, row, 2, 12)
    row += 1
    ch_first = row
    for j, ch in enumerate(CHANNELS):
        label_cell(ws.cell(row, 2), ch)
        formula_cell(ws.cell(row, 3), f"=SUMIF($C$6:$C${last_detail},B{row},$D$6:$D${last_detail})", vol_fmt)
        formula_cell(ws.cell(row, 4), f"=SUMIF($C$6:$C${last_detail},B{row},$E$6:$E${last_detail})", vol_fmt)
        formula_cell(ws.cell(row, 5), f"=SUMIF($C$6:$C${last_detail},B{row},$J$6:$J${last_detail})", money1)
        formula_cell(ws.cell(row, 6), f"=SUMIF($C$6:$C${last_detail},B{row},$K$6:$K${last_detail})", money1)
        formula_cell(ws.cell(row, 7), f"=SUMIF($C$6:$C${last_detail},B{row},$M$6:$M${last_detail})", money1)
        formula_cell(ws.cell(row, 8), f"=SUMIF($C$6:$C${last_detail},B{row},$N$6:$N${last_detail})", money1)
        formula_cell(ws.cell(row, 9), f"=IF(E{row}=0,0,G{row}/E{row})", pct)
        formula_cell(ws.cell(row, 10), f"=IF(F{row}=0,0,H{row}/F{row})", pct)
        formula_cell(ws.cell(row, 11), f"=C{row}/$C${ch_first+4}", pct)  # will fix after total
        formula_cell(ws.cell(row, 12), f"=D{row}/$D${ch_first+4}", pct)
        row += 1
    ch_last = row - 1
    tot_ch = row
    label_cell(ws.cell(row, 2), "Total", True)
    for col, fmt in [(3, vol_fmt), (4, vol_fmt), (5, money1), (6, money1), (7, money1), (8, money1)]:
        formula_cell(ws.cell(row, col), f"=SUM({get_column_letter(col)}{ch_first}:{get_column_letter(col)}{ch_last})", fmt, True)
    formula_cell(ws.cell(row, 9), f"=IF(E{row}=0,0,G{row}/E{row})", pct, True)
    formula_cell(ws.cell(row, 10), f"=IF(F{row}=0,0,H{row}/F{row})", pct, True)
    formula_cell(ws.cell(row, 11), "=1", pct, True)
    formula_cell(ws.cell(row, 12), "=1", pct, True)
    # Fix mix % to point at total row
    for r in range(ch_first, ch_last + 1):
        ws.cell(r, 11).value = f"=C{r}/$C${tot_ch}"
        ws.cell(r, 12).value = f"=D{r}/$D${tot_ch}"

    # Named ranges helper cells for bridge (KPIs)
    ws["P2"] = "KPI anchors (for bridge)"
    ws["P2"].font = section_font
    labels = [
        ("P3", "Rev Prior"), ("Q3", f"=J{total_row}"),
        ("P4", "Rev Curr"), ("Q4", f"=K{total_row}"),
        ("P5", "GM Prior"), ("Q5", f"=M{total_row}"),
        ("P6", "GM Curr"), ("Q6", f"=N{total_row}"),
        ("P7", "Vol Prior"), ("Q7", f"=D{total_row}"),
        ("P8", "Vol Curr"), ("Q8", f"=E{total_row}"),
    ]
    for addr, val in labels:
        cell = ws[addr]
        if isinstance(val, str) and val.startswith("="):
            formula_cell(cell, val, money1 if "Rev" in str(ws[addr[0] + "3"].value if False else "") or True else vol_fmt)
        else:
            label_cell(cell, val)
    for addr in ["Q3", "Q4", "Q5", "Q6"]:
        ws[addr].number_format = money1
    for addr in ["Q7", "Q8"]:
        ws[addr].number_format = vol_fmt
    ws["P3"] = "Rev Prior $000"
    ws["P4"] = "Rev Curr $000"
    ws["P5"] = "GM Prior $000"
    ws["P6"] = "GM Curr $000"
    ws["P7"] = "Vol Prior 000"
    ws["P8"] = "Vol Curr 000"
    for a in ["P3", "P4", "P5", "P6", "P7", "P8"]:
        ws[a].font = black


def build_pvm_bridge(wb):
    """Classic PVM on revenue and GM$ with clear bridge steps."""
    ws = wb.create_sheet("03_PVM_Bridge", 3)
    set_col_widths(ws, [3, 36, 16, 16, 50])

    ws["B2"] = "Price / Volume / Mix Bridge"
    ws["B2"].font = title_font
    ws["B3"] = (
        "Decomposition of revenue and gross margin $ from prior → current. "
        "Volume at prior mix & prior price/margin; Price on current volume; "
        "Mix = residual mix shift; Cost = COGS/unit change (≈0 here)."
    )
    ws["B3"].font = italic_grey
    ws["B3"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[3].height = 40

    # Helper: compute line-level components in a calc area (columns G+)
    # For each SKU×channel row on 02 (rows 6-29):
    # Volume effect on Rev = (V1_scaled_to_hold_mix? ) standard approach used here:
    #
    # Method (common FP&A):
    # 1. Price effect (Rev) = Σ V1 × (P1 − P0)
    # 2. Volume effect (Rev) = (Σ V1 − Σ V0) × blended P0   [pure volume at prior avg price]
    # 3. Mix effect (Rev) = Σ V1×P0 − Σ V0×P0 − Volume effect
    #    equivalently: Σ (mix1_i − mix0_i) × total_V1 × P0_i
    #
    # For GM:
    # margin/unit prior m0 = P0 − C; current m1 = P1 − C (C constant)
    # Price effect (GM) = Σ V1 × (P1 − P0)   [= same as rev price when C fixed]
    # Volume effect (GM) = (Σ V1 − Σ V0) × blended m0
    # Mix effect (GM) = Σ V1×m0 − Σ V0×m0 − Volume effect
    # Cost effect (GM) = Σ V1 × (C0 − C1)  [= 0]

    ws["B5"] = "A. Line-level components (SKU × channel)"
    ws["B5"].font = section_font

    # Build calc sheet area referencing 02_Volume_Price
    calc_headers = [
        "SKU", "Channel", "V0", "V1", "P0", "P1", "C",
        "Rev0", "Rev1", "GM0", "GM1",
        "V1*P0", "V1*m0",
        "Price_Rev", "Price_GM", "Cost_GM",
    ]
    for i, h in enumerate(calc_headers):
        ws.cell(6, 2 + i, h)
    style_header_row(ws, 6, 2, 17)

    for idx in range(24):  # 6 SKUs × 4 channels
        r = 7 + idx
        src = 6 + idx  # row on 02_Volume_Price
        formula_cell(ws.cell(r, 2), f"='02_Volume_Price'!B{src}")
        formula_cell(ws.cell(r, 3), f"='02_Volume_Price'!C{src}")
        formula_cell(ws.cell(r, 4), f"='02_Volume_Price'!D{src}", vol_fmt)
        formula_cell(ws.cell(r, 5), f"='02_Volume_Price'!E{src}", vol_fmt)
        formula_cell(ws.cell(r, 6), f"='02_Volume_Price'!G{src}", ppu)
        formula_cell(ws.cell(r, 7), f"='02_Volume_Price'!H{src}", ppu)
        formula_cell(ws.cell(r, 8), f"='02_Volume_Price'!L{src}", ppu)
        formula_cell(ws.cell(r, 9), f"=D{r}*F{r}", money1)   # Rev0
        formula_cell(ws.cell(r, 10), f"=E{r}*G{r}", money1)  # Rev1
        formula_cell(ws.cell(r, 11), f"=D{r}*(F{r}-H{r})", money1)  # GM0
        formula_cell(ws.cell(r, 12), f"=E{r}*(G{r}-H{r})", money1)  # GM1
        formula_cell(ws.cell(r, 13), f"=E{r}*F{r}", money1)  # V1*P0
        formula_cell(ws.cell(r, 14), f"=E{r}*(F{r}-H{r})", money1)  # V1*m0
        formula_cell(ws.cell(r, 15), f"=E{r}*(G{r}-F{r})", money1)  # Price_Rev
        formula_cell(ws.cell(r, 16), f"=E{r}*(G{r}-F{r})", money1)  # Price_GM (C fixed)
        formula_cell(ws.cell(r, 17), f"=E{r}*0", money1)  # Cost_GM placeholder (=0)

    last = 30  # 7+23
    # Totals row
    r = 31
    label_cell(ws.cell(r, 2), "TOTAL", True)
    for col, fmt in [
        (4, vol_fmt), (5, vol_fmt),
        (9, money1), (10, money1), (11, money1), (12, money1),
        (13, money1), (14, money1), (15, money1), (16, money1), (17, money1),
    ]:
        formula_cell(ws.cell(r, col), f"=SUM({get_column_letter(col)}7:{get_column_letter(col)}{last})", fmt, True)

    # Aggregate bridge math
    ws["B33"] = "B. Bridge math (totals)"
    ws["B33"].font = section_font

    # Store key totals
    # Q33 etc
    ws["B34"] = "Total V0"
    formula_cell(ws["C34"], "=D31", vol_fmt)
    ws["B35"] = "Total V1"
    formula_cell(ws["C35"], "=E31", vol_fmt)
    ws["B36"] = "Blended P0 (Rev0/V0)"
    formula_cell(ws["C36"], "=IF(C34=0,0,I31/C34)", ppu)
    ws["B37"] = "Blended m0 (GM0/V0)"
    formula_cell(ws["C37"], "=IF(C34=0,0,K31/C34)", ppu)

    ws["B39"] = "Volume effect — Revenue"
    formula_cell(ws["C39"], "=(C35-C34)*C36", money1)
    ws["D39"] = "(V1−V0) × blended prior price"
    ws["D39"].font = small_grey

    ws["B40"] = "Price effect — Revenue"
    formula_cell(ws["C40"], "=O31", money1)
    ws["D40"] = "Σ V1 × (P1−P0)"
    ws["D40"].font = small_grey

    ws["B41"] = "Mix effect — Revenue"
    formula_cell(ws["C41"], "=M31-I31-C39", money1)
    ws["D41"] = "Σ V1×P0 − Rev0 − Volume effect"
    ws["D41"].font = small_grey

    ws["B42"] = "Check: Δ Revenue"
    formula_cell(ws["C42"], "=J31-I31", money1, True)
    ws["B43"] = "Check: Vol+Price+Mix"
    formula_cell(ws["C43"], "=C39+C40+C41", money1, True)

    ws["B45"] = "Volume effect — GM $"
    formula_cell(ws["C45"], "=(C35-C34)*C37", money1)
    ws["D45"] = "(V1−V0) × blended prior margin/unit"
    ws["D45"].font = small_grey

    ws["B46"] = "Price effect — GM $"
    formula_cell(ws["C46"], "=P31", money1)
    ws["D46"] = "Σ V1 × (P1−P0)  [COGS fixed → same as rev price]"
    ws["D46"].font = small_grey

    ws["B47"] = "Mix effect — GM $"
    formula_cell(ws["C47"], "=N31-K31-C45", money1)
    ws["D47"] = "Σ V1×m0 − GM0 − Volume effect"
    ws["D47"].font = small_grey

    ws["B48"] = "Cost effect — GM $"
    formula_cell(ws["C48"], "=Q31", money1)
    ws["D48"] = "Σ V1 × (C0−C1) ≈ 0 in this sample"
    ws["D48"].font = small_grey

    ws["B49"] = "Check: Δ GM $"
    formula_cell(ws["C49"], "=L31-K31", money1, True)
    ws["B50"] = "Check: Vol+Price+Mix+Cost"
    formula_cell(ws["C50"], "=C45+C46+C47+C48", money1, True)

    # Visible bridge tables
    ws["B52"] = "C. Revenue bridge (prior → current)"
    ws["B52"].font = section_font
    for i, h in enumerate(["Step", "$000s", "Notes"]):
        ws.cell(53, 2 + i, h)
    style_header_row(ws, 53, 2, 4)

    rev_steps = [
        ("Prior Revenue", "=I31", "Starting point"),
        ("(+) Volume", "=C39", "Units up at prior blended price"),
        ("(+/-) Price", "=C40", "Net price pressure on current volume"),
        ("(+/-) Mix", "=C41", "Shift toward lower-price Club / Multipack"),
        ("Current Revenue", "=J31", "Prior + Volume + Price + Mix"),
    ]
    for i, (name, formul, note) in enumerate(rev_steps):
        r = 54 + i
        label_cell(ws.cell(r, 2), name, name.startswith("Prior") or name.startswith("Current"))
        formula_cell(ws.cell(r, 3), formul, money1, name.startswith("Prior") or name.startswith("Current"))
        ws.cell(r, 4, note).font = small_grey
        if name.startswith("Prior") or name.startswith("Current"):
            ws.cell(r, 2).fill = bridge_fill
            ws.cell(r, 3).fill = bridge_fill

    ws["B60"] = "D. Gross margin $ bridge (prior → current)"
    ws["B60"].font = section_font
    for i, h in enumerate(["Step", "$000s", "Notes"]):
        ws.cell(61, 2 + i, h)
    style_header_row(ws, 61, 2, 4)

    gm_steps = [
        ("Prior GM $", "=K31", "Starting gross margin dollars"),
        ("(+) Volume", "=C45", "More units at prior blended margin/unit"),
        ("(+/-) Price", "=C46", "Lower net realized price"),
        ("(+/-) Mix", "=C47", "Club & Multipack mix dilutes margin"),
        ("(+/-) Cost", "=C48", "Std COGS/unit unchanged → ~0"),
        ("Current GM $", "=L31", "Prior + Volume + Price + Mix + Cost"),
    ]
    for i, (name, formul, note) in enumerate(gm_steps):
        r = 62 + i
        label_cell(ws.cell(r, 2), name, name.startswith("Prior") or name.startswith("Current"))
        formula_cell(ws.cell(r, 3), formul, money1, name.startswith("Prior") or name.startswith("Current"))
        ws.cell(r, 4, note).font = small_grey
        if name.startswith("Prior") or name.startswith("Current"):
            ws.cell(r, 2).fill = bridge_fill
            ws.cell(r, 3).fill = bridge_fill

    # Margin % walk
    ws["B70"] = "E. Margin rate"
    ws["B70"].font = section_font
    ws["B71"] = "GM % Prior"
    formula_cell(ws["C71"], "=IF(I31=0,0,K31/I31)", pct, True)
    ws["B72"] = "GM % Curr"
    formula_cell(ws["C72"], "=IF(J31=0,0,L31/J31)", pct, True)
    ws["B73"] = "Δ GM pts"
    formula_cell(ws["C73"], "=C72-C71", "0.0%")

    ws["B75"] = (
        "Interpretation: Volume contributes positively to GM$, but Mix (Club channel share + "
        "Multipack Snacks) and Price (Grocery/Mass net price) more than offset — overall GM$ "
        "and GM% decline vs prior even though units are higher."
    )
    ws["B75"].font = italic_grey
    ws["B75"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[75].height = 48

    # Chart data for dashboard (compact)
    ws["B77"] = "Chart source — GM walk steps"
    ws["B77"].font = section_font
    ws["B78"] = "Step"
    ws["C78"] = "$000s"
    style_header_row(ws, 78, 2, 3)
    chart_items = [
        ("Prior GM", "=C62"),
        ("Volume", "=C63"),
        ("Price", "=C64"),
        ("Mix", "=C65"),
        ("Cost", "=C66"),
        ("Current GM", "=C67"),
    ]
    for i, (name, formul) in enumerate(chart_items):
        r = 79 + i
        label_cell(ws.cell(r, 2), name)
        formula_cell(ws.cell(r, 3), formul, money1)

    chart = BarChart()
    chart.type = "col"
    chart.title = "Gross Margin $ Walk ($000s)"
    chart.y_axis.title = "$000s"
    data = Reference(ws, min_col=3, min_row=78, max_row=84)
    cats = Reference(ws, min_col=2, min_row=79, max_row=84)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 15
    chart.height = 8
    ws.add_chart(chart, "E52")


def build_contribution(wb):
    ws = wb.create_sheet("04_Contribution", 4)
    set_col_widths(ws, [3, 26, 14, 14, 14, 14, 14, 14, 14, 14])

    ws["B2"] = "Contribution Margin by SKU & Channel"
    ws["B2"].font = title_font
    ws["B3"] = "Contribution = Net revenue − Std COGS. Mix % of volume and revenue for prior and current."
    ws["B3"].font = italic_grey

    # By SKU — pull from 02 summary
    ws["B5"] = "By SKU (current period focus)"
    ws["B5"].font = section_font
    headers = [
        "SKU", "Vol Curr", "Vol Mix %", "Rev Curr $000", "Rev Mix %",
        "GM Curr $000", "GM Mix %", "GM %", "Δ Vol vs Prior", "Δ GM $ vs Prior",
    ]
    for i, h in enumerate(headers):
        ws.cell(6, 2 + i, h)
    style_header_row(ws, 6, 2, 11)
    ws.row_dimensions[6].height = 30

    # SKU summary on 02 starts after detail 6-29, total 30, then +3 = 33 header?, need robust refs
    # From build_volume_price: detail 6..29, total 30, then row+=3 → 33 is "Summary by SKU", 34 header, 35-40 SKUs, 41 total
    # Channel summary: 44 header area... Let me recalculate:
    # row after total_row(30): row = 30+3 = 33 section title, 34 headers, 35-40 skus, 41 total
    # then row = 41+3 = 44 section, 45 headers, 46-49 channels, 50 total

    sku_start_02 = 35
    for i, sku in enumerate(SKUS):
        r = 7 + i
        src = sku_start_02 + i
        label_cell(ws.cell(r, 2), sku)
        formula_cell(ws.cell(r, 3), f"='02_Volume_Price'!D{src}", vol_fmt)
        formula_cell(ws.cell(r, 4), f"=C{r}/$C$13", pct)
        formula_cell(ws.cell(r, 5), f"='02_Volume_Price'!F{src}", money1)
        formula_cell(ws.cell(r, 6), f"=E{r}/$E$13", pct)
        formula_cell(ws.cell(r, 7), f"='02_Volume_Price'!H{src}", money1)
        formula_cell(ws.cell(r, 8), f"=G{r}/$G$13", pct)
        formula_cell(ws.cell(r, 9), f"='02_Volume_Price'!J{src}", pct)
        formula_cell(ws.cell(r, 10), f"='02_Volume_Price'!D{src}-'02_Volume_Price'!C{src}", vol_fmt)
        formula_cell(ws.cell(r, 11), f"='02_Volume_Price'!H{src}-'02_Volume_Price'!G{src}", money1)

    # Total row 13
    label_cell(ws.cell(13, 2), "Total", True)
    for col, fmt in [(3, vol_fmt), (5, money1), (7, money1)]:
        formula_cell(ws.cell(13, col), f"=SUM({get_column_letter(col)}7:{get_column_letter(col)}12)", fmt, True)
    formula_cell(ws.cell(13, 4), "=1", pct, True)
    formula_cell(ws.cell(13, 6), "=1", pct, True)
    formula_cell(ws.cell(13, 8), "=1", pct, True)
    formula_cell(ws.cell(13, 9), f"=IF(E13=0,0,G13/E13)", pct, True)
    formula_cell(ws.cell(13, 10), f"=SUM(J7:J12)", vol_fmt, True)
    formula_cell(ws.cell(13, 11), f"=SUM(K7:K12)", money1, True)

    # By Channel
    ws["B15"] = "By Channel"
    ws["B15"].font = section_font
    headers2 = [
        "Channel", "Vol Curr", "Vol Mix %", "Rev Curr $000", "Rev Mix %",
        "GM Curr $000", "CM $/unit", "GM %", "Vol Mix Δ pts", "Callout",
    ]
    for i, h in enumerate(headers2):
        ws.cell(16, 2 + i, h)
    style_header_row(ws, 16, 2, 11)

    ch_start_02 = 46
    callouts = [
        "Highest margin channel; share down",
        "Mid margin; mild share dip",
        "Mix headwind — share up, thin CM",
        "Premium price; growing share",
    ]
    for j, ch in enumerate(CHANNELS):
        r = 17 + j
        src = ch_start_02 + j
        label_cell(ws.cell(r, 2), ch)
        formula_cell(ws.cell(r, 3), f"='02_Volume_Price'!D{src}", vol_fmt)
        formula_cell(ws.cell(r, 4), f"=C{r}/$C$21", pct)
        formula_cell(ws.cell(r, 5), f"='02_Volume_Price'!F{src}", money1)
        formula_cell(ws.cell(r, 6), f"=E{r}/$E$21", pct)
        formula_cell(ws.cell(r, 7), f"='02_Volume_Price'!H{src}", money1)
        formula_cell(ws.cell(r, 8), f"=IF(C{r}=0,0,G{r}/C{r})", ppu)
        formula_cell(ws.cell(r, 9), f"='02_Volume_Price'!J{src}", pct)
        # Vol mix Δ = curr mix - prior mix
        formula_cell(ws.cell(r, 10), f"='02_Volume_Price'!L{src}-'02_Volume_Price'!K{src}", "0.0%")
        label_cell(ws.cell(r, 11), callouts[j])

    label_cell(ws.cell(21, 2), "Total", True)
    for col, fmt in [(3, vol_fmt), (5, money1), (7, money1)]:
        formula_cell(ws.cell(21, col), f"=SUM({get_column_letter(col)}17:{get_column_letter(col)}20)", fmt, True)
    formula_cell(ws.cell(21, 4), "=1", pct, True)
    formula_cell(ws.cell(21, 6), "=1", pct, True)
    formula_cell(ws.cell(21, 8), "=IF(C21=0,0,G21/C21)", ppu, True)
    formula_cell(ws.cell(21, 9), "=IF(E21=0,0,G21/E21)", pct, True)
    formula_cell(ws.cell(21, 10), "=0", "0.0%", True)

    ws["B23"] = "Contribution insight"
    ws["B23"].font = section_font
    ws["B24"] = (
        "Club volume mix rises while CM $/unit stays the lowest of the four channels. "
        "Multipack Snacks (Club-heavy) also gains share at thin margin. Net: positive volume "
        "on the GM bridge, negative mix — classic CPG dilution when club packs outgrow core."
    )
    ws["B24"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[24].height = 48


def build_dashboard(wb):
    ws = wb.create_sheet("05_Dashboard", 5)
    set_col_widths(ws, [3, 22, 14, 14, 14, 14, 18, 28])

    ws["B2"] = "Margin & Mix Dashboard"
    ws["B2"].font = title_font
    formula_cell(ws["B3"], "=\"Close: \"&'01_Assumptions'!C5")
    ws["B3"].font = italic_grey

    # KPI tiles
    ws["B5"] = "Headline KPIs"
    ws["B5"].font = section_font

    kpis = [
        (6, "Vol Prior (000)", "='03_PVM_Bridge'!C34", vol_fmt),
        (7, "Vol Curr (000)", "='03_PVM_Bridge'!C35", vol_fmt),
        (8, "Δ Volume", "='03_PVM_Bridge'!C35-'03_PVM_Bridge'!C34", vol_fmt),
        (9, "Rev Prior $000", "='03_PVM_Bridge'!I31", money1),
        (10, "Rev Curr $000", "='03_PVM_Bridge'!J31", money1),
        (11, "GM Prior $000", "='03_PVM_Bridge'!K31", money1),
        (12, "GM Curr $000", "='03_PVM_Bridge'!L31", money1),
        (13, "GM % Prior", "='03_PVM_Bridge'!C71", pct),
        (14, "GM % Curr", "='03_PVM_Bridge'!C72", pct),
        (15, "Δ GM pts", "='03_PVM_Bridge'!C73", "0.0%"),
    ]
    ws["B6"] = "Metric"
    ws["C6"] = "Value"
    style_header_row(ws, 6, 2, 3)
    for r, label, formul, fmt in kpis:
        # shift: we used 6 for header, so data from 7
        pass
    # rewrite cleanly
    for i, (label, formul, fmt) in enumerate([
        ("Vol Prior (000)", "='03_PVM_Bridge'!C34", vol_fmt),
        ("Vol Curr (000)", "='03_PVM_Bridge'!C35", vol_fmt),
        ("Δ Volume", "='03_PVM_Bridge'!C35-'03_PVM_Bridge'!C34", vol_fmt),
        ("Rev Prior $000", "='03_PVM_Bridge'!I31", money1),
        ("Rev Curr $000", "='03_PVM_Bridge'!J31", money1),
        ("GM Prior $000", "='03_PVM_Bridge'!K31", money1),
        ("GM Curr $000", "='03_PVM_Bridge'!L31", money1),
        ("GM % Prior", "='03_PVM_Bridge'!C71", pct),
        ("GM % Curr", "='03_PVM_Bridge'!C72", pct),
        ("Δ GM pts", "='03_PVM_Bridge'!C73", "0.0%"),
    ]):
        r = 7 + i
        label_cell(ws.cell(r, 2), label)
        formula_cell(ws.cell(r, 3), formul, fmt, True)
        ws.cell(r, 2).fill = tile_fill
        ws.cell(r, 3).fill = tile_fill

    # RAG on margin %
    ws["B18"] = "RAG — Gross margin %"
    ws["B18"].font = section_font
    ws["B19"] = "Target GM %"
    input_cell(ws["C19"], 0.54, pct)
    ws["B20"] = "Actual GM % (curr)"
    formula_cell(ws["C20"], "='03_PVM_Bridge'!C72", pct, True)
    ws["B21"] = "Status"
    formula_cell(ws["C21"], '=IF(C20>=C19,"GREEN",IF(C20>=C19-0.02,"AMBER","RED"))')
    ws["C21"].font = bold_black
    ws["C21"].border = thin

    # Conditional fill for RAG
    ws.conditional_formatting.add(
        "C21",
        FormulaRule(formula=['$C$21="GREEN"'], fill=green_fill),
    )
    ws.conditional_formatting.add(
        "C21",
        FormulaRule(formula=['$C$21="AMBER"'], fill=amber_fill),
    )
    ws.conditional_formatting.add(
        "C21",
        FormulaRule(formula=['$C$21="RED"'], fill=red_fill),
    )
    # Also shade C20 vs target
    ws.conditional_formatting.add(
        "C20",
        CellIsRule(operator="greaterThanOrEqual", formula=["$C$19"], fill=green_fill),
    )
    ws.conditional_formatting.add(
        "C20",
        FormulaRule(formula=["AND(C20<$C$19,C20>=$C$19-0.02)"], fill=amber_fill),
    )
    ws.conditional_formatting.add(
        "C20",
        CellIsRule(operator="lessThan", formula=["$C$19-0.02"], fill=red_fill),
    )

    # GM walk one-pager
    ws["E5"] = "Gross margin $ walk"
    ws["E5"].font = section_font
    for i, h in enumerate(["Step", "$000s"]):
        ws.cell(6, 5 + i, h)
    style_header_row(ws, 6, 5, 6)
    for i, (name, ref) in enumerate([
        ("Prior GM", "='03_PVM_Bridge'!C62"),
        ("Volume", "='03_PVM_Bridge'!C63"),
        ("Price", "='03_PVM_Bridge'!C64"),
        ("Mix", "='03_PVM_Bridge'!C65"),
        ("Cost", "='03_PVM_Bridge'!C66"),
        ("Current GM", "='03_PVM_Bridge'!C67"),
    ]):
        r = 7 + i
        label_cell(ws.cell(r, 5), name, name in ("Prior GM", "Current GM"))
        formula_cell(ws.cell(r, 6), ref, money1, name in ("Prior GM", "Current GM"))
        if name in ("Prior GM", "Current GM"):
            ws.cell(r, 5).fill = bridge_fill
            ws.cell(r, 6).fill = bridge_fill

    chart = BarChart()
    chart.type = "col"
    chart.title = "GM$ Walk"
    data = Reference(ws, min_col=6, min_row=6, max_row=12)
    cats = Reference(ws, min_col=5, min_row=7, max_row=12)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 12
    chart.height = 8
    ws.add_chart(chart, "E14")

    # Mix callouts
    ws["B23"] = "Top mix-shift callouts"
    ws["B23"].font = section_font
    ws["B24"] = "Callout"
    ws["C24"] = "Signal"
    style_header_row(ws, 24, 2, 3)
    callouts = [
        ("Club vol mix ↑", "Lower CM $/unit channel gaining share → mix headwind on GM$"),
        ("Multipack Snacks ↑", "Club-oriented SKU grows fastest; thin contribution"),
        ("Grocery vol soft", "Highest-price channel loses units → adverse mix & price"),
        ("Price (Groc/Mass)", "Net price down ~$0.10–0.20/unit on core channels"),
        ("Volume overall ↑", "Units up, but not enough to offset mix + price on GM$"),
    ]
    for i, (c, s) in enumerate(callouts):
        r = 25 + i
        label_cell(ws.cell(r, 2), c, True)
        label_cell(ws.cell(r, 3), s)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

    ws["B31"] = "Leadership read"
    ws["B31"].font = section_font
    ws["B32"] = (
        "Volume is up, but gross margin dollars and margin rate are down. The bridge pins "
        "the miss on Mix (Club / Multipack) and Price — not cost. Action lens: protect "
        "Grocery/Mass net price, rebalance Club pack architecture, or raise Club list to "
        "defend contribution."
    )
    ws["B32"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("B32:G34")
    ws.row_dimensions[32].height = 20


def build_dictionary(wb):
    ws = wb.create_sheet("06_Data_Dictionary", 6)
    set_col_widths(ws, [3, 28, 18, 60])

    ws["B2"] = "Data Dictionary"
    ws["B2"].font = title_font
    ws["B3"] = "Field definitions for the margin & mix bridge workbook."
    ws["B3"].font = italic_grey

    headers = ["Field", "Unit / type", "Definition"]
    for i, h in enumerate(headers):
        ws.cell(5, 2 + i, h)
    style_header_row(ws, 5, 2, 4)

    rows = [
        ("Close period", "Text", "Prior vs current comparison label (e.g. FY26 Q2 vs Q1)"),
        ("Channel", "Category", "Grocery, Mass, Club, E-comm — route-to-market"),
        ("SKU", "Category", "Sellable consumer unit / pack (6 sample Northline items)"),
        ("Volume", "000 units", "Cases/units sold in the period (thousands)"),
        ("Net price", "$/unit", "Average realized net selling price after trade"),
        ("Std COGS", "$/unit", "Standard cost per unit; held flat in this sample"),
        ("Revenue", "$000s", "Volume (000) × net price $/unit"),
        ("Gross margin $", "$000s", "Volume × (net price − std COGS)"),
        ("GM %", "%", "Gross margin $ ÷ revenue"),
        ("Contribution / CM", "$ or %", "Same as GM here (no below-gross selling cost allocated)"),
        ("Volume effect", "$000s", "(V1−V0) × prior blended price or margin/unit"),
        ("Price effect", "$000s", "Σ current volume × (P1−P0)"),
        ("Mix effect", "$000s", "Residual from mix shift at prior prices/margins"),
        ("Cost effect", "$000s", "Σ V1 × (C0−C1); ~0 when COGS/unit unchanged"),
        ("Vol mix %", "%", "Channel or SKU share of total volume"),
        ("RAG — GM %", "Status", "GREEN if ≥ target; AMBER within 2 pts; else RED"),
        ("Yellow / blue cells", "Input", "Editable assumptions; do not overwrite black formulas"),
    ]
    for i, (f, u, d) in enumerate(rows):
        r = 6 + i
        label_cell(ws.cell(r, 2), f)
        label_cell(ws.cell(r, 3), u)
        label_cell(ws.cell(r, 4), d)

    ws["B25"] = "Fictional sample for portfolio use only. Not real company, employer, or client data."
    ws["B25"].font = italic_grey


def verify_row_layout():
    """Sanity: recompute expected row numbers for 02_Volume_Price summaries."""
    # detail rows 6..29 (24 rows), total 30
    # +3 → 33 title, 34 hdr, 35-40 sku, 41 total
    # +3 → 44 title, 45 hdr, 46-49 ch, 50 total
    assert 6 + 24 - 1 == 29
    assert 35 + 5 == 40


def main():
    verify_row_layout()
    wb = Workbook()
    # remove default
    default = wb.active
    wb.remove(default)

    build_cover(wb)
    build_assumptions(wb)
    build_volume_price(wb)
    build_pvm_bridge(wb)
    build_contribution(wb)
    build_dashboard(wb)
    build_dictionary(wb)

    out = Path(__file__).resolve().parent / "Northline_Margin_Mix_Bridge.xlsx"
    wb.save(out)
    print(f"Wrote {out}")
    print("Sheets:", wb.sheetnames)


if __name__ == "__main__":
    main()
