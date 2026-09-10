#!/usr/bin/env python3
"""Build Sales Advanced Excel portfolio workbook (Tableau Sample Superstore).

Author attribution: GitHub shanitripathi417-code (no invented legal name).
Data: Tableau Public Sample Superstore (fictitious retail orders, USD).
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path("/workspace/excel-sales-advanced")
OUT = ROOT / "Sales_Advanced_Excel_Dashboard.xlsx"
DATA_XLS = ROOT / "data" / "sample_superstore.xls"

# ---------- palette ----------
NAVY = "1B3A4B"
TEAL = "0D7377"
GOLD = "C9A227"
LIGHT = "F5F7FA"
WHITE = "FFFFFF"
GREEN = "1B7A4E"
RED = "B91C1C"
SOFT_GREEN = "D1FAE5"
SOFT_RED = "FEE2E2"
SOFT_BLUE = "DBEAFE"
SOFT_AMBER = "FEF3C7"
GRAY = "64748B"
DARK = "0F172A"

thin = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
)

header_fill = PatternFill("solid", fgColor=NAVY)
header_font = Font(name="Calibri", bold=True, color=WHITE, size=11)
title_font = Font(name="Calibri", bold=True, color=NAVY, size=18)
section_font = Font(name="Calibri", bold=True, color=TEAL, size=13)
kpi_label = Font(name="Calibri", bold=True, color=GRAY, size=10)
kpi_value = Font(name="Calibri", bold=True, color=DARK, size=16)
body_font = Font(name="Calibri", size=10)
formula_font = Font(name="Consolas", size=9, color="1E3A5F")
label_font = Font(name="Calibri", bold=True, size=10, color=DARK)

usd_fmt = '"$"#,##0'
usd_fmt_dec = '"$"#,##0.00'
pct_fmt = "0.0%"
num_fmt = "#,##0"
date_fmt = "DD-MMM-YYYY"

# Portfolio target margins (stretch targets for Margin-vs-Target demos)
TARGET_MARGIN = {
    "Furniture": 0.08,
    "Office Supplies": 0.18,
    "Technology": 0.18,
}
SEGMENT_PRIORITY = {
    "Corporate": "P1-Strategic",
    "Home Office": "P2-Growth",
    "Consumer": "P3-Volume",
}
HIGH_VALUE = 1000  # USD


def style_header(ws, row, start_col, end_col):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin


def autosize(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def load_superstore():
    orders = pd.read_excel(DATA_XLS, sheet_name="Orders")
    people = pd.read_excel(DATA_XLS, sheet_name="People")
    # Map People: Regional Manager, Region
    managers = {r["Region"]: r["Regional Manager"] for _, r in people.iterrows()}

    rows = []
    for _, r in orders.iterrows():
        qty = int(r["Quantity"])
        sales = float(r["Sales"])
        discount = float(r["Discount"])
        profit = float(r["Profit"])
        cost = round(sales - profit, 4)
        if qty > 0 and discount < 1:
            unit = round(sales / qty / (1 - discount), 4)
        elif qty > 0:
            unit = round(sales / qty, 4)
        else:
            unit = 0.0
        rows.append(
            {
                "OrderID": str(r["Order ID"]),
                "OrderDate": r["Order Date"].to_pydatetime().date()
                if hasattr(r["Order Date"], "to_pydatetime")
                else r["Order Date"],
                "Region": str(r["Region"]),
                "ProductCategory": str(r["Category"]),
                "Product": str(r["Product Name"])[:80],
                "CustomerSegment": str(r["Segment"]),
                "Customer": str(r["Customer Name"]),
                "Quantity": qty,
                "UnitPrice": unit,
                "Discount": discount,
                "Revenue": round(sales, 4),
                "Cost": cost,
                "Profit": round(profit, 4),
                "State": str(r["State/Province"]),
                "SubCategory": str(r["Sub-Category"]),
            }
        )
    rows.sort(key=lambda x: (x["OrderDate"], x["OrderID"]))
    regions = sorted(managers.keys())
    categories = sorted(TARGET_MARGIN.keys())
    segments = sorted(SEGMENT_PRIORITY.keys())
    years = sorted({r["OrderDate"].year for r in rows})
    return rows, managers, regions, categories, segments, years


def build():
    rows, managers, REGIONS, CATEGORIES, SEGMENTS, YEARS = load_superstore()
    N_ROWS = len(rows)
    last = N_ROWS + 1
    year_min, year_max = min(YEARS), max(YEARS)
    yoy_new, yoy_old = year_max, year_max - 1

    wb = Workbook()

    # ===================== Instructions =====================
    ws = wb.active
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False
    autosize(ws, {c: 22 for c in "ABCDEFGHI"})
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 88

    ws["B2"] = "Sales Advanced Excel — Portfolio Project"
    ws["B2"].font = title_font
    ws["B3"] = (
        f"GitHub: shanitripathi417-code  ·  Advanced Excel / MIS portfolio  ·  "
        f"Currency: USD  ·  Source: Tableau Sample Superstore"
    )
    ws["B3"].font = Font(name="Calibri", italic=True, color=GRAY, size=11)
    ws["B4"] = "Open in Microsoft Excel (desktop or Excel 365). Formulas recalculate live."
    ws["B4"].font = Font(name="Calibri", size=10, color=TEAL)

    tips = [
        ("Sheet Map", ""),
        (
            "Raw_Data",
            f"{N_ROWS:,} retail order lines from Tableau Sample Superstore "
            f"({year_min}–{year_max}). Source fact table — values only.",
        ),
        (
            "Lookups",
            "Dimension tables: Region→Regional Manager (from Superstore People sheet), "
            "Category→Target Margin, Segment→Priority.",
        ),
        (
            "Clean_Model",
            "Derived columns with REAL formulas (Year, Month, Margin %, Net, XLOOKUP Manager, flags).",
        ),
        (
            "Advanced_Formulas",
            "Interview showcase: SUMIFS, COUNTIFS, AVERAGEIFS, INDEX/MATCH, XLOOKUP, IFS, TEXT, "
            "EOMONTH, RANK, SUMPRODUCT, TEXTJOIN, running totals, YoY.",
        ),
        (
            "Dashboard",
            "Executive KPIs + charts (Region, Category, Monthly trend, Segment mix). USD formats.",
        ),
        ("", ""),
        ("How to demo (interview)", ""),
        (
            "1. Story",
            "Start on Dashboard → point to KPIs → drill into charts → jump to Advanced_Formulas to prove formula fluency.",
        ),
        (
            "2. Lookups",
            "Show Clean_Model!RegionManager = XLOOKUP from Lookups (fallback INDEX/MATCH noted).",
        ),
        (
            "3. Conditionals",
            "Margin % conditional formatting: green ≥20%, red <10% (on Clean_Model & Dashboard).",
        ),
        ("4. IFS / flags", "High-value / loss / high-discount flags driven by nested IF / IFS."),
        (
            "5. Excel 365",
            "FILTER / UNIQUE / SORT blocks are optional dynamic-array demos — open in Excel 365 to spill.",
        ),
        ("", ""),
        ("Talking points", ""),
        (
            "Data model",
            "Star-style: fact (Raw_Data) + dims (Lookups) → Clean_Model → Dashboard aggregates.",
        ),
        (
            "Why formulas",
            "Every derived metric is a cell formula so interviewers can audit / change inputs.",
        ),
        (
            "Currency",
            "USD — Superstore is a U.S. retail sample published by Tableau Public.",
        ),
        (
            "Scale",
            f"{N_ROWS:,} order lines after clean — honest count from the public dataset (no downsampling).",
        ),
        (
            "Attribution",
            "Tableau Sample Superstore (Orders + People). Fictitious company data for demos; "
            "download: https://public.tableau.com/app/sample-data/sample_-_superstore.xls",
        ),
        ("", ""),
        (
            "Constraints noted",
            "Built with Python (openpyxl). No Excel Desktop on build machine; formulas are Excel-native.",
        ),
    ]
    r = 6
    for a, b in tips:
        if a and not b:
            ws.cell(row=r, column=2, value=a).font = section_font
        elif a:
            ws.cell(row=r, column=2, value=a).font = label_font
            ws.cell(row=r, column=3, value=b).font = body_font
            ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True)
        r += 1
    ws.row_dimensions[2].height = 28

    # ===================== Lookups =====================
    ws = wb.create_sheet("Lookups")
    ws.sheet_view.showGridLines = False
    autosize(ws, {"A": 3, "B": 14, "C": 22, "D": 3, "E": 18, "F": 14, "G": 3, "H": 14, "I": 16})

    ws["B2"] = "Lookup / Dimension Tables"
    ws["B2"].font = title_font
    ws["B3"] = "Used by XLOOKUP / VLOOKUP / INDEX-MATCH demos on Clean_Model & Advanced_Formulas"
    ws["B3"].font = Font(name="Calibri", italic=True, color=GRAY, size=10)

    ws["B5"] = "Region → Manager"
    ws["B5"].font = section_font
    ws["B6"] = "Region"
    ws["C6"] = "Manager"
    style_header(ws, 6, 2, 3)
    for i, reg in enumerate(REGIONS, 7):
        ws.cell(row=i, column=2, value=reg).border = thin
        ws.cell(row=i, column=3, value=managers[reg]).border = thin
        ws.cell(row=i, column=2).font = body_font
        ws.cell(row=i, column=3).font = body_font
    region_last = 6 + len(REGIONS)  # last data row for region lookup

    ws["E5"] = "Category → Target Margin"
    ws["E5"].font = section_font
    ws["E6"] = "Category"
    ws["F6"] = "TargetMargin"
    style_header(ws, 6, 5, 6)
    for i, cat in enumerate(CATEGORIES, 7):
        ws.cell(row=i, column=5, value=cat).border = thin
        cell = ws.cell(row=i, column=6, value=TARGET_MARGIN[cat])
        cell.number_format = pct_fmt
        cell.border = thin
        ws.cell(row=i, column=5).font = body_font
        cell.font = body_font
    cat_last = 6 + len(CATEGORIES)

    ws["H5"] = "Segment → Priority"
    ws["H5"].font = section_font
    ws["H6"] = "Segment"
    ws["I6"] = "Priority"
    style_header(ws, 6, 8, 9)
    for i, seg in enumerate(SEGMENTS, 7):
        ws.cell(row=i, column=8, value=seg).border = thin
        ws.cell(row=i, column=9, value=SEGMENT_PRIORITY[seg]).border = thin
        ws.cell(row=i, column=8).font = body_font
        ws.cell(row=i, column=9).font = body_font
    seg_last = 6 + len(SEGMENTS)

    ws["B14"] = (
        f"Managers sourced from Tableau Superstore People sheet. "
        f"Target margins are portfolio stretch targets for Margin-vs-Target demos "
        f"(Furniture actual ~2.6%; Office Supplies / Technology ~17%)."
    )
    ws["B14"].font = Font(name="Calibri", size=9, color=GRAY)

    # ===================== Raw_Data =====================
    ws = wb.create_sheet("Raw_Data", 1)
    headers = [
        "OrderID",
        "OrderDate",
        "Region",
        "ProductCategory",
        "Product",
        "CustomerSegment",
        "Customer",
        "Quantity",
        "UnitPrice",
        "Discount",
        "Revenue",
        "Cost",
        "Profit",
        "State",
        "SubCategory",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, 1, len(headers))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:O{N_ROWS + 1}"

    for i, row in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=row["OrderID"]).font = body_font
        cell = ws.cell(row=i, column=2, value=row["OrderDate"])
        cell.number_format = date_fmt
        cell.font = body_font
        for col, key in enumerate(
            ["Region", "ProductCategory", "Product", "CustomerSegment", "Customer"], 3
        ):
            ws.cell(row=i, column=col, value=row[key]).font = body_font
        ws.cell(row=i, column=8, value=row["Quantity"]).font = body_font
        cell = ws.cell(row=i, column=9, value=row["UnitPrice"])
        cell.number_format = usd_fmt_dec
        cell.font = body_font
        cell = ws.cell(row=i, column=10, value=row["Discount"])
        cell.number_format = pct_fmt
        cell.font = body_font
        for col, key in [(11, "Revenue"), (12, "Cost"), (13, "Profit")]:
            cell = ws.cell(row=i, column=col, value=row[key])
            cell.number_format = usd_fmt
            cell.font = body_font
        ws.cell(row=i, column=14, value=row["State"]).font = body_font
        ws.cell(row=i, column=15, value=row["SubCategory"]).font = body_font

    autosize(
        ws,
        {
            "A": 16,
            "B": 13,
            "C": 10,
            "D": 16,
            "E": 36,
            "F": 13,
            "G": 18,
            "H": 9,
            "I": 11,
            "J": 10,
            "K": 11,
            "L": 11,
            "M": 11,
            "N": 14,
            "O": 12,
        },
    )

    # ===================== Clean_Model =====================
    ws = wb.create_sheet("Clean_Model", 2)
    cm_headers = [
        "OrderID",
        "OrderDate",
        "Year",
        "Month",
        "MonthName",
        "Region",
        "ProductCategory",
        "Product",
        "CustomerSegment",
        "Customer",
        "Quantity",
        "UnitPrice",
        "Discount",
        "Gross",
        "NetRevenue",
        "Cost",
        "Profit",
        "ProfitMargin%",
        "RegionManager",
        "TargetMargin",
        "MarginVsTarget",
        "Priority",
        "HighValueFlag",
        "LossFlag",
        "HighDiscountFlag",
    ]
    for c, h in enumerate(cm_headers, 1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, 1, len(cm_headers))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:Y{last}"

    for r in range(2, last + 1):
        ws.cell(row=r, column=1, value=f"=Raw_Data!A{r}")
        cell = ws.cell(row=r, column=2, value=f"=Raw_Data!B{r}")
        cell.number_format = date_fmt
        ws.cell(row=r, column=3, value=f"=YEAR(B{r})")
        ws.cell(row=r, column=4, value=f"=MONTH(B{r})")
        ws.cell(row=r, column=5, value=f'=TEXT(B{r},"MMM-YYYY")')
        ws.cell(row=r, column=6, value=f"=Raw_Data!C{r}")
        ws.cell(row=r, column=7, value=f"=Raw_Data!D{r}")
        ws.cell(row=r, column=8, value=f"=Raw_Data!E{r}")
        ws.cell(row=r, column=9, value=f"=Raw_Data!F{r}")
        ws.cell(row=r, column=10, value=f"=Raw_Data!G{r}")
        ws.cell(row=r, column=11, value=f"=Raw_Data!H{r}")
        cell = ws.cell(row=r, column=12, value=f"=Raw_Data!I{r}")
        cell.number_format = usd_fmt_dec
        cell = ws.cell(row=r, column=13, value=f"=Raw_Data!J{r}")
        cell.number_format = pct_fmt
        cell = ws.cell(row=r, column=14, value=f"=K{r}*L{r}")
        cell.number_format = usd_fmt
        cell = ws.cell(row=r, column=15, value=f"=N{r}*(1-M{r})")
        cell.number_format = usd_fmt
        cell = ws.cell(row=r, column=16, value=f"=Raw_Data!L{r}")
        cell.number_format = usd_fmt
        cell = ws.cell(row=r, column=17, value=f"=O{r}-P{r}")
        cell.number_format = usd_fmt
        cell = ws.cell(row=r, column=18, value=f"=IF(O{r}=0,0,Q{r}/O{r})")
        cell.number_format = pct_fmt
        ws.cell(
            row=r,
            column=19,
            value=f'=XLOOKUP(F{r},Lookups!$B$7:$B${region_last},Lookups!$C$7:$C${region_last},"Unknown")',
        )
        cell = ws.cell(
            row=r,
            column=20,
            value=f"=XLOOKUP(G{r},Lookups!$E$7:$E${cat_last},Lookups!$F$7:$F${cat_last},0)",
        )
        cell.number_format = pct_fmt
        cell = ws.cell(row=r, column=21, value=f"=R{r}-T{r}")
        cell.number_format = "0.0%;[Red]-0.0%"
        ws.cell(
            row=r,
            column=22,
            value=f'=XLOOKUP(I{r},Lookups!$H$7:$H${seg_last},Lookups!$I$7:$I${seg_last},"")',
        )
        ws.cell(row=r, column=23, value=f'=IF(O{r}>={HIGH_VALUE},"HIGH VALUE","")')
        ws.cell(row=r, column=24, value=f'=IF(Q{r}<0,"LOSS","")')
        ws.cell(
            row=r,
            column=25,
            value=f'=IFS(M{r}>=0.15,"Deep Discount",M{r}>=0.1,"Moderate",TRUE,"Standard")',
        )

    ws.conditional_formatting.add(
        f"R2:R{last}",
        CellIsRule(
            operator="greaterThanOrEqual",
            formula=["0.2"],
            fill=PatternFill("solid", fgColor=SOFT_GREEN),
            font=Font(color=GREEN, bold=True),
        ),
    )
    ws.conditional_formatting.add(
        f"R2:R{last}",
        CellIsRule(
            operator="lessThan",
            formula=["0.1"],
            fill=PatternFill("solid", fgColor=SOFT_RED),
            font=Font(color=RED, bold=True),
        ),
    )
    ws.conditional_formatting.add(
        f"U2:U{last}",
        CellIsRule(
            operator="greaterThanOrEqual",
            formula=["0"],
            fill=PatternFill("solid", fgColor=SOFT_GREEN),
        ),
    )
    ws.conditional_formatting.add(
        f"U2:U{last}",
        CellIsRule(
            operator="lessThan",
            formula=["0"],
            fill=PatternFill("solid", fgColor=SOFT_RED),
        ),
    )

    autosize(
        ws,
        {
            "A": 16,
            "B": 12,
            "C": 7,
            "D": 7,
            "E": 11,
            "F": 10,
            "G": 15,
            "H": 28,
            "I": 13,
            "J": 16,
            "K": 9,
            "L": 11,
            "M": 9,
            "N": 11,
            "O": 12,
            "P": 11,
            "Q": 11,
            "R": 12,
            "S": 18,
            "T": 12,
            "U": 13,
            "V": 13,
            "W": 13,
            "X": 10,
            "Y": 14,
        },
    )
    note_row = last + 2
    ws.cell(
        row=note_row,
        column=1,
        value=(
            "NOTE: Columns C–Y are Excel formulas. XLOOKUP needs Excel 365/2021; "
            "INDEX/MATCH equivalent shown on Advanced_Formulas. "
            "NetRevenue ≈ Raw Sales (list price recovered from Sales÷Qty÷(1−Discount))."
        ),
    ).font = Font(name="Calibri", italic=True, size=9, color=GRAY)

    # ===================== Advanced_Formulas =====================
    ws = wb.create_sheet("Advanced_Formulas", 3)
    ws.sheet_view.showGridLines = False
    autosize(
        ws,
        {
            "A": 3,
            "B": 32,
            "C": 18,
            "D": 58,
            "E": 14,
            "F": 14,
            "G": 14,
            "H": 14,
            "I": 14,
            "J": 14,
            "K": 18,
            "L": 18,
        },
    )

    ws["B2"] = "Advanced Formulas Showcase"
    ws["B2"].font = title_font
    ws["B3"] = (
        "Labels + live formulas against Clean_Model / Raw_Data / Lookups  ·  Interview walkthrough panel"
    )
    ws["B3"].font = Font(name="Calibri", italic=True, color=GRAY, size=10)

    default_region = "West" if "West" in REGIONS else REGIONS[0]
    default_cat = "Technology" if "Technology" in CATEGORIES else CATEGORIES[0]
    default_year = yoy_new if yoy_new in YEARS else YEARS[-1]

    ws["B5"] = "Parameters (change these)"
    ws["B5"].font = section_font
    ws["B6"] = "Selected Region"
    ws["C6"] = default_region
    ws["C6"].fill = PatternFill("solid", fgColor=SOFT_AMBER)
    ws["C6"].border = thin
    ws["C6"].font = Font(bold=True)
    ws["B7"] = "Selected Category"
    ws["C7"] = default_cat
    ws["C7"].fill = PatternFill("solid", fgColor=SOFT_AMBER)
    ws["C7"].border = thin
    ws["C7"].font = Font(bold=True)
    ws["B8"] = "Selected Year"
    ws["C8"] = default_year
    ws["C8"].fill = PatternFill("solid", fgColor=SOFT_AMBER)
    ws["C8"].border = thin
    ws["C8"].font = Font(bold=True)
    ws["B9"] = "High-Value Threshold ($)"
    ws["C9"] = HIGH_VALUE
    ws["C9"].number_format = usd_fmt
    ws["C9"].fill = PatternFill("solid", fgColor=SOFT_AMBER)
    ws["C9"].border = thin

    dv1 = DataValidation(
        type="list", formula1='"' + ",".join(REGIONS) + '"', allow_blank=False
    )
    dv2 = DataValidation(
        type="list", formula1='"' + ",".join(CATEGORIES) + '"', allow_blank=False
    )
    dv3 = DataValidation(
        type="list", formula1='"' + ",".join(str(y) for y in YEARS) + '"', allow_blank=False
    )
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    ws.add_data_validation(dv3)
    dv1.add(ws["C6"])
    dv2.add(ws["C7"])
    dv3.add(ws["C8"])

    ws["B11"] = "Function / Pattern"
    ws["C11"] = "Result"
    ws["D11"] = "Formula (audit)"
    style_header(ws, 11, 2, 4)

    cm = "Clean_Model"
    demos = [
        (
            "SUMIFS — Revenue by Region",
            f"=SUMIFS({cm}!O:O,{cm}!F:F,$C$6)",
            "=SUMIFS(Clean_Model!O:O,Clean_Model!F:F,$C$6)",
            usd_fmt,
        ),
        (
            "SUMIFS — Rev Region+Category",
            f"=SUMIFS({cm}!O:O,{cm}!F:F,$C$6,{cm}!G:G,$C$7)",
            "=SUMIFS(Clean_Model!O:O,Clean_Model!F:F,$C$6,Clean_Model!G:G,$C$7)",
            usd_fmt,
        ),
        (
            "COUNTIFS — Orders (Region+Year)",
            f"=COUNTIFS({cm}!F:F,$C$6,{cm}!C:C,$C$8)",
            "=COUNTIFS(Clean_Model!F:F,$C$6,Clean_Model!C:C,$C$8)",
            num_fmt,
        ),
        (
            "AVERAGEIFS — AOV by Segment",
            f'=AVERAGEIFS({cm}!O:O,{cm}!I:I,"Corporate")',
            '=AVERAGEIFS(Clean_Model!O:O,Clean_Model!I:I,"Corporate")',
            usd_fmt,
        ),
        (
            "XLOOKUP — Manager for Region",
            f'=XLOOKUP($C$6,Lookups!$B$7:$B${region_last},Lookups!$C$7:$C${region_last},"N/A")',
            f"=XLOOKUP($C$6,Lookups!$B$7:$B${region_last},Lookups!$C$7:$C${region_last},\"N/A\")",
            None,
        ),
        (
            "INDEX/MATCH — Target Margin",
            f"=INDEX(Lookups!$F$7:$F${cat_last},MATCH($C$7,Lookups!$E$7:$E${cat_last},0))",
            f"=INDEX(Lookups!$F$7:$F${cat_last},MATCH($C$7,Lookups!$E$7:$E${cat_last},0))",
            pct_fmt,
        ),
        (
            "IFS — Region tier label",
            '=IFS($C$6="West","Tier-A",$C$6="East","Tier-A",$C$6="Central","Tier-B",TRUE,"Tier-C")',
            '=IFS($C$6="West","Tier-A",$C$6="East","Tier-A",$C$6="Central","Tier-B",TRUE,"Tier-C")',
            None,
        ),
        (
            "Nested IF — performance band",
            f'=IF(SUMIFS({cm}!O:O,{cm}!F:F,$C$6)>=600000,"Star",IF(SUMIFS({cm}!O:O,{cm}!F:F,$C$6)>=400000,"Solid","Watch"))',
            '=IF(SUMIFS(...Region)>=600k,"Star",IF(...>=400k,"Solid","Watch"))',
            None,
        ),
        (
            "TEXT — formatted period label",
            '=TEXT(DATE($C$8,1,1),"YYYY")&" · Region "&$C$6',
            '=TEXT(DATE($C$8,1,1),"YYYY")&" · Region "&$C$6',
            None,
        ),
        (
            "EOMONTH — month-end for Jan of year",
            "=EOMONTH(DATE($C$8,1,1),0)",
            "=EOMONTH(DATE($C$8,1,1),0)",
            date_fmt,
        ),
        (
            "DATE — calendar year start",
            "=DATE($C$8,1,1)",
            "=DATE($C$8,1,1)  // calendar year start",
            date_fmt,
        ),
        (
            "SUMPRODUCT — unique customers (approx)",
            f"=SUMPRODUCT(({cm}!F$2:F${last}=$C$6)/COUNTIFS({cm}!J$2:J${last},{cm}!J$2:J${last},{cm}!F$2:F${last},$C$6))",
            "=SUMPRODUCT((Region=sel)/COUNTIFS(Customer,Customer,Region,sel))",
            "0.0",
        ),
        (
            "SUMPRODUCT — weighted margin",
            f"=SUMPRODUCT(({cm}!F$2:F${last}=$C$6)*({cm}!Q$2:Q${last}))/SUMPRODUCT(({cm}!F$2:F${last}=$C$6)*({cm}!O$2:O${last}))",
            "=SUMPRODUCT((Region=sel)*Profit)/SUMPRODUCT((Region=sel)*NetRev)",
            pct_fmt,
        ),
        (
            "COUNTIFS — High value orders",
            f'=COUNTIFS({cm}!O:O,">="&$C$9,{cm}!F:F,$C$6)',
            '=COUNTIFS(Clean_Model!O:O,">="&$C$9,Clean_Model!F:F,$C$6)',
            num_fmt,
        ),
        (
            "AVERAGEIFS — Margin % filter",
            f"=AVERAGEIFS({cm}!R:R,{cm}!G:G,$C$7,{cm}!C:C,$C$8)",
            "=AVERAGEIFS(Clean_Model!R:R,Clean_Model!G:G,$C$7,Clean_Model!C:C,$C$8)",
            pct_fmt,
        ),
        (
            "MAXIFS — Peak order (Region)",
            f"=MAXIFS({cm}!O:O,{cm}!F:F,$C$6)",
            "=MAXIFS(Clean_Model!O:O,Clean_Model!F:F,$C$6)",
            usd_fmt,
        ),
        (
            "MINIFS — Smallest order (Cat)",
            f"=MINIFS({cm}!O:O,{cm}!G:G,$C$7)",
            "=MINIFS(Clean_Model!O:O,Clean_Model!G:G,$C$7)",
            usd_fmt,
        ),
        (
            "TEXTJOIN — region list",
            f'=TEXTJOIN(" | ",TRUE,Lookups!B7:B{region_last})',
            f'=TEXTJOIN(" | ",TRUE,Lookups!B7:B{region_last})',
            None,
        ),
        (
            f"YoY — {yoy_new} vs {yoy_old} Revenue",
            f"=SUMIFS({cm}!O:O,{cm}!C:C,{yoy_new})/SUMIFS({cm}!O:O,{cm}!C:C,{yoy_old})-1",
            f"=SUMIFS(Net,Year,{yoy_new})/SUMIFS(Net,Year,{yoy_old})-1",
            "0.0%;[Red]-0.0%",
        ),
        (
            "IFERROR — safe divide demo",
            f"=IFERROR(SUMIFS({cm}!Q:Q,{cm}!F:F,$C$6)/SUMIFS({cm}!O:O,{cm}!F:F,$C$6),0)",
            "=IFERROR(Profit/Revenue for region, 0)",
            pct_fmt,
        ),
    ]

    for i, (label, formula, audit, fmt) in enumerate(demos, 12):
        ws.cell(row=i, column=2, value=label).font = label_font
        ws.cell(row=i, column=2).border = thin
        cell = ws.cell(row=i, column=3, value=formula)
        cell.font = Font(name="Calibri", bold=True, size=11, color=DARK)
        cell.border = thin
        cell.fill = PatternFill("solid", fgColor=SOFT_BLUE)
        if fmt:
            cell.number_format = fmt
        ws.cell(row=i, column=4, value=audit).font = formula_font
        ws.cell(row=i, column=4).border = thin

    last_demo = 11 + len(demos)

    r0 = last_demo + 2
    ws.cell(row=r0, column=2, value="RANK panel — Regions by Revenue").font = section_font
    ws.cell(row=r0 + 1, column=2, value="Region").font = header_font
    ws.cell(row=r0 + 1, column=2).fill = header_fill
    ws.cell(row=r0 + 1, column=3, value="Revenue").font = header_font
    ws.cell(row=r0 + 1, column=3).fill = header_fill
    ws.cell(row=r0 + 1, column=4, value="Rank").font = header_font
    ws.cell(row=r0 + 1, column=4).fill = header_fill

    for i, reg in enumerate(REGIONS):
        rr = r0 + 2 + i
        ws.cell(row=rr, column=2, value=reg).border = thin
        cell = ws.cell(row=rr, column=3, value=f"=SUMIF({cm}!F:F,B{rr},{cm}!O:O)")
        cell.number_format = usd_fmt
        cell.border = thin
        cell = ws.cell(
            row=rr,
            column=4,
            value=f"=RANK(C{rr},$C${r0+2}:$C${r0+1+len(REGIONS)},0)",
        )
        cell.border = thin
        cell.alignment = Alignment(horizontal="center")

    rt = r0 + 2 + len(REGIONS) + 2
    ws.cell(
        row=rt, column=2, value="Running Total — Monthly Net Revenue (Selected Year)"
    ).font = section_font
    ws.cell(row=rt + 1, column=2, value="Month").font = header_font
    ws.cell(row=rt + 1, column=2).fill = header_fill
    ws.cell(row=rt + 1, column=3, value="Month Rev").font = header_font
    ws.cell(row=rt + 1, column=3).fill = header_fill
    ws.cell(row=rt + 1, column=4, value="Running Total").font = header_font
    ws.cell(row=rt + 1, column=4).fill = header_fill

    for m in range(1, 13):
        rr = rt + 1 + m
        ws.cell(row=rr, column=2, value=m).border = thin
        cell = ws.cell(
            row=rr,
            column=3,
            value=f"=SUMIFS({cm}!O:O,{cm}!C:C,$C$8,{cm}!D:D,B{rr})",
        )
        cell.number_format = usd_fmt
        cell.border = thin
        cell = ws.cell(row=rr, column=4, value=f"=SUM($C${rt+2}:C{rr})")
        cell.number_format = usd_fmt
        cell.border = thin

    da = rt + 16
    ws.cell(row=da, column=2, value="Optional — Excel 365 Dynamic Arrays").font = section_font
    ws.cell(
        row=da + 1,
        column=2,
        value=(
            "Open in Excel 365 for spill results. openpyxl writes the formulas; "
            "older Excel shows #NAME? / needs Ctrl+Shift+Enter era workarounds."
        ),
    ).font = Font(name="Calibri", italic=True, size=9, color=GRAY)

    ws.cell(row=da + 3, column=2, value="UNIQUE regions").font = label_font
    ws.cell(row=da + 3, column=3, value=f"=UNIQUE({cm}!F2:F{last})")
    ws.cell(row=da + 3, column=3).font = formula_font
    ws.cell(row=da + 3, column=4, value=f"=UNIQUE(Clean_Model!F2:F{last})").font = formula_font

    ws.cell(row=da + 4, column=2, value="SORT categories").font = label_font
    ws.cell(row=da + 4, column=3, value=f"=SORT(UNIQUE(Clean_Model!G2:G{last}))")
    ws.cell(row=da + 4, column=3).font = formula_font
    ws.cell(row=da + 4, column=4, value="Dynamic array — Excel 365").font = Font(
        italic=True, size=9, color=GRAY
    )

    ws.cell(row=da + 5, column=2, value="FILTER high-value (param region)").font = label_font
    ws.cell(
        row=da + 5,
        column=3,
        value=f'=FILTER({cm}!A2:A{last},({cm}!F2:F{last}=$C$6)*({cm}!O2:O{last}>=$C$9),"None")',
    )
    ws.cell(row=da + 5, column=3).font = formula_font
    ws.cell(row=da + 5, column=4, value="FILTER OrderIDs where Region=param & Net>=threshold").font = Font(
        italic=True, size=9, color=GRAY
    )

    ws.cell(row=da + 7, column=2, value="VLOOKUP classic (Manager)").font = label_font
    ws.cell(
        row=da + 7,
        column=3,
        value=f"=VLOOKUP($C$6,Lookups!$B$7:$C${region_last},2,FALSE)",
    )
    ws.cell(row=da + 7, column=3).font = formula_font
    ws.cell(row=da + 7, column=3).fill = PatternFill("solid", fgColor=SOFT_BLUE)
    ws.cell(row=da + 7, column=3).border = thin

    # ===================== Dashboard =====================
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    for col in range(1, 16):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 3
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["J"].width = 14
    ws.column_dimensions["K"].width = 14

    ws.merge_cells("B2:K2")
    ws["B2"] = "Sales Performance Dashboard"
    ws["B2"].font = Font(name="Calibri", bold=True, color=WHITE, size=22)
    ws["B2"].fill = PatternFill("solid", fgColor=NAVY)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    for col in range(2, 12):
        ws.cell(row=2, column=col).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[2].height = 36

    ws.merge_cells("B3:K3")
    ws["B3"] = (
        f"GitHub: shanitripathi417-code  ·  Advanced Excel Portfolio  ·  "
        f"Currency: USD  ·  Period: {year_min}–{year_max}  ·  Tableau Sample Superstore"
    )
    ws["B3"].font = Font(name="Calibri", color=WHITE, size=10)
    ws["B3"].fill = PatternFill("solid", fgColor=TEAL)
    for col in range(2, 12):
        ws.cell(row=3, column=col).fill = PatternFill("solid", fgColor=TEAL)

    kpis = [
        ("B", "Total Revenue", f"=SUM({cm}!O:O)", usd_fmt),
        ("D", "Total Profit", f"=SUM({cm}!Q:Q)", usd_fmt),
        ("F", "Profit Margin %", f"=IFERROR(SUM({cm}!Q:Q)/SUM({cm}!O:O),0)", pct_fmt),
        ("H", "Order Lines", f"=COUNTA({cm}!A2:A{last})", num_fmt),
        ("J", "AOV (line)", f"=IFERROR(SUM({cm}!O:O)/COUNTA({cm}!A2:A{last}),0)", usd_fmt),
    ]
    for col_letter, label, formula, fmt in kpis:
        col = ord(col_letter) - 64
        cell = ws.cell(row=5, column=col, value=label)
        cell.font = kpi_label
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=5, column=col + 1).fill = PatternFill("solid", fgColor=LIGHT)
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
        cell = ws.cell(row=6, column=col, value=formula)
        cell.font = kpi_value
        cell.number_format = fmt
        cell.fill = PatternFill("solid", fgColor=WHITE)
        ws.merge_cells(start_row=6, start_column=col, end_row=7, end_column=col + 1)
        ws.cell(row=6, column=col).alignment = Alignment(horizontal="center", vertical="center")
        for rr in (6, 7):
            for cc in (col, col + 1):
                ws.cell(row=rr, column=cc).border = Border(
                    left=Side(style="medium", color=TEAL),
                    right=Side(style="medium", color=TEAL),
                    top=Side(style="medium", color=TEAL),
                    bottom=Side(style="medium", color=TEAL),
                )
                ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=WHITE)
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[7].height = 22

    ws.conditional_formatting.add(
        "F6",
        CellIsRule(
            operator="greaterThanOrEqual",
            formula=["0.15"],
            fill=PatternFill("solid", fgColor=SOFT_GREEN),
            font=Font(name="Calibri", bold=True, color=GREEN, size=16),
        ),
    )
    ws.conditional_formatting.add(
        "F6",
        CellIsRule(
            operator="lessThan",
            formula=["0.10"],
            fill=PatternFill("solid", fgColor=SOFT_RED),
            font=Font(name="Calibri", bold=True, color=RED, size=16),
        ),
    )

    # Region summary
    ws["M5"] = "Region"
    ws["N5"] = "Revenue"
    ws["O5"] = "Profit"
    style_header(ws, 5, 13, 15)
    for i, reg in enumerate(REGIONS):
        rr = 6 + i
        ws.cell(row=rr, column=13, value=reg)
        cell = ws.cell(row=rr, column=14, value=f"=SUMIF({cm}!F:F,M{rr},{cm}!O:O)")
        cell.number_format = usd_fmt
        cell = ws.cell(row=rr, column=15, value=f"=SUMIF({cm}!F:F,M{rr},{cm}!Q:Q)")
        cell.number_format = usd_fmt
    region_chart_end = 5 + len(REGIONS)

    # Category summary
    cat_header_row = region_chart_end + 2
    ws.cell(row=cat_header_row, column=13, value="Category")
    ws.cell(row=cat_header_row, column=14, value="Revenue")
    style_header(ws, cat_header_row, 13, 14)
    for i, cat in enumerate(CATEGORIES):
        rr = cat_header_row + 1 + i
        ws.cell(row=rr, column=13, value=cat)
        cell = ws.cell(row=rr, column=14, value=f"=SUMIF({cm}!G:G,M{rr},{cm}!O:O)")
        cell.number_format = usd_fmt
    cat_chart_end = cat_header_row + len(CATEGORIES)

    # Monthly trend — full year span
    month_header = cat_chart_end + 2
    ws.cell(row=month_header, column=13, value="Period")
    ws.cell(row=month_header, column=14, value="Revenue")
    style_header(ws, month_header, 13, 14)
    months = []
    for y in range(year_min, year_max + 1):
        for m in range(1, 13):
            months.append((y, m))
    for i, (year, month) in enumerate(months):
        rr = month_header + 1 + i
        ws.cell(row=rr, column=13, value=date(year, month, 1))
        ws.cell(row=rr, column=13).number_format = "MMM-YY"
        cell = ws.cell(
            row=rr,
            column=14,
            value=f"=SUMIFS({cm}!O:O,{cm}!C:C,{year},{cm}!D:D,{month})",
        )
        cell.number_format = usd_fmt
    month_end = month_header + len(months)

    # Segment mix
    seg_header = month_end + 2
    ws.cell(row=seg_header, column=13, value="Segment")
    ws.cell(row=seg_header, column=14, value="Revenue")
    style_header(ws, seg_header, 13, 14)
    for i, seg in enumerate(SEGMENTS):
        rr = seg_header + 1 + i
        ws.cell(row=rr, column=13, value=seg)
        cell = ws.cell(row=rr, column=14, value=f"=SUMIF({cm}!I:I,M{rr},{cm}!O:O)")
        cell.number_format = usd_fmt
    seg_end = seg_header + len(SEGMENTS)

    # Charts
    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "Revenue by Region"
    chart1.style = 10
    chart1.y_axis.title = "Revenue ($)"
    data = Reference(ws, min_col=14, min_row=5, max_row=region_chart_end)
    cats = Reference(ws, min_col=13, min_row=6, max_row=region_chart_end)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.width = 12
    chart1.height = 8
    ws.add_chart(chart1, "B9")

    chart2 = BarChart()
    chart2.type = "bar"
    chart2.title = "Revenue by Category"
    chart2.style = 10
    data = Reference(ws, min_col=14, min_row=cat_header_row, max_row=cat_chart_end)
    cats = Reference(ws, min_col=13, min_row=cat_header_row + 1, max_row=cat_chart_end)
    chart2.add_data(data, titles_from_data=True)
    chart2.set_categories(cats)
    chart2.width = 12
    chart2.height = 8
    ws.add_chart(chart2, "H9")

    chart3 = LineChart()
    chart3.title = "Monthly Revenue Trend"
    chart3.style = 10
    chart3.y_axis.title = "Revenue ($)"
    data = Reference(ws, min_col=14, min_row=month_header, max_row=month_end)
    cats = Reference(ws, min_col=13, min_row=month_header + 1, max_row=month_end)
    chart3.add_data(data, titles_from_data=True)
    chart3.set_categories(cats)
    chart3.width = 18
    chart3.height = 8
    ws.add_chart(chart3, "B25")

    chart4 = PieChart()
    chart4.title = "Segment Mix (Revenue)"
    labels = Reference(ws, min_col=13, min_row=seg_header + 1, max_row=seg_end)
    data = Reference(ws, min_col=14, min_row=seg_header, max_row=seg_end)
    chart4.add_data(data, titles_from_data=True)
    chart4.set_categories(labels)
    chart4.dataLabels = DataLabelList()
    chart4.dataLabels.showPercent = True
    chart4.dataLabels.showVal = False
    chart4.dataLabels.showCatName = False
    chart4.width = 10
    chart4.height = 8
    ws.add_chart(chart4, "H25")

    ws["B35"] = (
        "Source: Clean_Model formulas ← Raw_Data (Tableau Sample Superstore)  ·  "
        "Change parameters on Advanced_Formulas  ·  Chart data tables in columns M–O"
    )
    ws["B35"].font = Font(name="Calibri", size=8, color=GRAY, italic=True)

    ws.column_dimensions["M"].width = 16
    ws.column_dimensions["N"].width = 14
    ws.column_dimensions["O"].width = 12

    order = [
        "Instructions",
        "Dashboard",
        "Raw_Data",
        "Lookups",
        "Clean_Model",
        "Advanced_Formulas",
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.save(OUT)
    meta = {
        "rows": N_ROWS,
        "years": YEARS,
        "regions": REGIONS,
        "categories": CATEGORIES,
        "segments": SEGMENTS,
        "managers": managers,
        "total_revenue": sum(r["Revenue"] for r in rows),
        "total_profit": sum(r["Profit"] for r in rows),
        "yoy_new": yoy_new,
        "yoy_old": yoy_old,
    }
    print(f"Saved: {OUT}")
    print(f"Rows: {N_ROWS}")
    print(f"Sheets: {wb.sheetnames}")
    print(f"Size KB: {OUT.stat().st_size / 1024:.1f}")
    print(f"Years: {YEARS}")
    print(f"Revenue: {meta['total_revenue']:.2f} Profit: {meta['total_profit']:.2f}")
    return meta


if __name__ == "__main__":
    build()
