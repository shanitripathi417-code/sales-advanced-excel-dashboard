#!/usr/bin/env python3
"""Render portfolio screenshots (Excel-like UI mockups) from Superstore aggregates."""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import FancyBboxPatch, Rectangle
from PIL import Image, ImageDraw, ImageFont

ROOT = Path("/workspace/excel-sales-advanced")
SHOTS = ROOT / "screenshots"
FRAMES = ROOT / "artifacts" / "frames"
DATA = ROOT / "data" / "sample_superstore.xls"
W, H = 1280, 720

NAVY = (27, 58, 75)
TEAL = (13, 115, 119)
GOLD = (201, 162, 39)
LIGHT = (245, 247, 250)
WHITE = (255, 255, 255)
DARK = (15, 23, 42)
GRAY = (100, 116, 139)
GREEN = (27, 122, 78)
SOFT_BLUE = (219, 234, 254)
SOFT_AMBER = (254, 243, 199)
GRID = (226, 232, 240)
HEADER = (27, 58, 75)

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_M = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"


def font(path, size):
    return ImageFont.truetype(path, size)


def money(n):
    if abs(n) >= 1_000_000:
        return f"${n/1_000_000:.2f}M"
    if abs(n) >= 1_000:
        return f"${n:,.0f}"
    return f"${n:,.2f}"


def load():
    orders = pd.read_excel(DATA, sheet_name="Orders")
    people = pd.read_excel(DATA, sheet_name="People")
    orders["Year"] = orders["Order Date"].dt.year
    orders["Month"] = orders["Order Date"].dt.month
    return orders, people


def draw_excel_chrome(draw, title_tab="Dashboard", sheet_tabs=None):
    # Window chrome
    draw.rectangle((0, 0, W, H), fill=(232, 236, 241))
    # Title bar
    draw.rectangle((0, 0, W, 36), fill=(45, 55, 72))
    draw.text((16, 10), "Sales_Advanced_Excel_Dashboard.xlsx — Excel", font=font(FONT, 14), fill=WHITE)
    # Ribbon strip
    draw.rectangle((0, 36, W, 68), fill=(250, 250, 252))
    draw.text((16, 44), "File   Home   Insert   Formulas   Data   View", font=font(FONT, 12), fill=GRAY)
    # Formula bar
    draw.rectangle((0, 68, W, 98), fill=WHITE)
    draw.rectangle((8, 72, 60, 94), fill=LIGHT)
    draw.text((18, 76), "fx", font=font(FONT_B, 13), fill=TEAL)
    draw.rectangle((68, 72, W - 12, 94), outline=GRID, width=1)
    draw.text((76, 76), title_tab, font=font(FONT_M, 12), fill=DARK)
    # Sheet area background
    draw.rectangle((0, 98, W, H - 28), fill=WHITE)
    # Sheet tabs
    if sheet_tabs is None:
        sheet_tabs = [
            "Instructions",
            "Dashboard",
            "Raw_Data",
            "Lookups",
            "Clean_Model",
            "Advanced_Formulas",
        ]
    x = 8
    for tab in sheet_tabs:
        active = tab == title_tab
        tw = int(draw.textlength(tab, font=font(FONT_B if active else FONT, 11)) + 18)
        y0 = H - 28
        if active:
            draw.rectangle((x, y0, x + tw, H), fill=WHITE)
            draw.rectangle((x, y0, x + tw, y0 + 3), fill=TEAL)
            draw.text((x + 9, y0 + 8), tab, font=font(FONT_B, 11), fill=TEAL)
        else:
            draw.rectangle((x, y0, x + tw, H), fill=(226, 232, 240))
            draw.text((x + 9, y0 + 8), tab, font=font(FONT, 11), fill=GRAY)
        x += tw + 4


def rounded(draw, xy, fill, radius=10, outline=None):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline)


def make_dashboard(orders: pd.DataFrame):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    draw_excel_chrome(d, "Dashboard")

    # Content area
    content_top = 110
    # Banner
    d.rectangle((16, content_top, W - 16, content_top + 44), fill=NAVY)
    d.text((28, content_top + 10), "Sales Performance Dashboard", font=font(FONT_B, 24), fill=WHITE)
    d.rectangle((16, content_top + 44, W - 16, content_top + 68), fill=TEAL)
    d.text(
        (28, content_top + 50),
        "GitHub: shanitripathi417-code  ·  Advanced Excel Portfolio  ·  Currency: USD  ·  Period: 2023–2026  ·  Tableau Sample Superstore",
        font=font(FONT, 11),
        fill=WHITE,
    )

    rev = orders["Sales"].sum()
    profit = orders["Profit"].sum()
    margin = profit / rev
    n = len(orders)
    aov = rev / n

    kpis = [
        ("Total Revenue", money(rev)),
        ("Total Profit", money(profit)),
        ("Profit Margin %", f"{margin*100:.1f}%"),
        ("Order Lines", f"{n:,}"),
        ("AOV (line)", money(aov)),
    ]
    kx = 24
    kw = 230
    for i, (lab, val) in enumerate(kpis):
        x0 = kx + i * (kw + 12)
        y0 = content_top + 84
        rounded(d, (x0, y0, x0 + kw, y0 + 70), WHITE, radius=6, outline=TEAL)
        d.rectangle((x0, y0, x0 + kw, y0 + 22), fill=LIGHT)
        d.text((x0 + 12, y0 + 4), lab, font=font(FONT_B, 11), fill=GRAY)
        color = GREEN if lab.startswith("Profit Margin") else DARK
        d.text((x0 + 12, y0 + 30), val, font=font(FONT_B, 22), fill=color)

    # Charts via matplotlib
    fig, axes = plt.subplots(2, 2, figsize=(11.8, 4.6), dpi=100)
    fig.patch.set_facecolor("white")

    by_reg = orders.groupby("Region")["Sales"].sum().reindex(["Central", "East", "South", "West"])
    ax = axes[0, 0]
    ax.bar(by_reg.index, by_reg.values / 1000, color="#0D7377")
    ax.set_title("Revenue by Region ($K)", fontsize=11, color="#1B3A4B", fontweight="bold")
    ax.set_ylabel("$K")
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(labelsize=9)

    by_cat = orders.groupby("Category")["Sales"].sum().sort_values()
    ax = axes[0, 1]
    ax.barh(by_cat.index, by_cat.values / 1000, color="#C9A227")
    ax.set_title("Revenue by Category ($K)", fontsize=11, color="#1B3A4B", fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(labelsize=9)

    monthly = (
        orders.assign(Period=orders["Order Date"].dt.to_period("M"))
        .groupby("Period")["Sales"]
        .sum()
        .sort_index()
    )
    ax = axes[1, 0]
    xs = range(len(monthly))
    ax.plot(xs, monthly.values / 1000, color="#0D7377", linewidth=2, marker="o", markersize=2.5)
    ax.fill_between(xs, monthly.values / 1000, alpha=0.15, color="#0D7377")
    ax.set_title("Monthly Revenue Trend ($K)", fontsize=11, color="#1B3A4B", fontweight="bold")
    step = max(1, len(monthly) // 8)
    ticks = list(xs)[::step]
    ax.set_xticks(ticks)
    ax.set_xticklabels([str(monthly.index[i]) for i in ticks], rotation=30, ha="right", fontsize=8)
    ax.spines[["top", "right"]].set_visible(False)

    by_seg = orders.groupby("Segment")["Sales"].sum()
    ax = axes[1, 1]
    colors = ["#1B3A4B", "#0D7377", "#C9A227"]
    ax.pie(
        by_seg.values,
        labels=by_seg.index,
        autopct="%1.0f%%",
        colors=colors[: len(by_seg)],
        textprops={"fontsize": 9},
    )
    ax.set_title("Segment Mix", fontsize=11, color="#1B3A4B", fontweight="bold")

    fig.tight_layout(pad=1.2)
    chart_path = SHOTS / "_chart_tmp.png"
    fig.savefig(chart_path, dpi=100, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    chart = Image.open(chart_path).convert("RGBA")
    # Fit charts under KPIs
    max_w = W - 40
    ratio = max_w / chart.width
    chart = chart.resize((max_w, int(chart.height * ratio)), Image.Resampling.LANCZOS)
    y_chart = content_top + 168
    if y_chart + chart.height > H - 36:
        chart = chart.resize((max_w, H - 36 - y_chart), Image.Resampling.LANCZOS)
    img.paste(chart, (20, y_chart), chart if chart.mode == "RGBA" else None)
    chart_path.unlink(missing_ok=True)

    out = SHOTS / "dashboard.png"
    img.save(out)
    print("wrote", out)
    return img


def make_raw_preview(orders: pd.DataFrame):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    draw_excel_chrome(d, "Raw_Data")

    d.text((20, 110), "Raw_Data — Tableau Sample Superstore (source fact table)", font=font(FONT_B, 16), fill=NAVY)
    d.text(
        (20, 136),
        f"{len(orders):,} order lines  ·  {orders['Order Date'].min().date()} → {orders['Order Date'].max().date()}  ·  USD",
        font=font(FONT, 12),
        fill=GRAY,
    )

    cols = [
        ("Order ID", 150),
        ("Order Date", 110),
        ("Region", 80),
        ("Category", 130),
        ("Segment", 110),
        ("Customer", 150),
        ("Qty", 50),
        ("Sales", 90),
        ("Profit", 90),
    ]
    x0, y0 = 16, 170
    # header
    x = x0
    for name, wcol in cols:
        d.rectangle((x, y0, x + wcol, y0 + 28), fill=HEADER)
        d.text((x + 6, y0 + 6), name, font=font(FONT_B, 11), fill=WHITE)
        x += wcol

    preview = orders.sort_values("Order Date").head(14)
    for i, (_, r) in enumerate(preview.iterrows()):
        y = y0 + 28 + i * 28
        bg = LIGHT if i % 2 == 0 else WHITE
        d.rectangle((x0, y, x0 + sum(c[1] for c in cols), y + 28), fill=bg)
        vals = [
            str(r["Order ID"]),
            r["Order Date"].strftime("%d-%b-%Y"),
            str(r["Region"]),
            str(r["Category"]),
            str(r["Segment"]),
            str(r["Customer Name"])[:18],
            str(int(r["Quantity"])),
            f"${r['Sales']:,.2f}",
            f"${r['Profit']:,.2f}",
        ]
        x = x0
        for (name, wcol), val in zip(cols, vals):
            d.rectangle((x, y, x + wcol, y + 28), outline=GRID)
            d.text((x + 6, y + 7), val, font=font(FONT, 11), fill=DARK)
            x += wcol

    d.text(
        (20, H - 56),
        "Columns mapped into workbook model (+ State, SubCategory). Derived Net/Cost/Margin live on Clean_Model.",
        font=font(FONT, 11),
        fill=GRAY,
    )
    out = SHOTS / "raw_data_preview.png"
    img.save(out)
    print("wrote", out)
    return img


def make_advanced(orders: pd.DataFrame, people: pd.DataFrame):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    draw_excel_chrome(d, "Advanced_Formulas")

    d.text((20, 110), "Advanced Formulas Showcase", font=font(FONT_B, 18), fill=NAVY)
    d.text(
        (20, 136),
        "Parameter-driven demos against Clean_Model  ·  change amber cells → results update in Excel",
        font=font(FONT, 12),
        fill=GRAY,
    )

    # Parameters box
    rounded(d, (20, 165, 420, 290), LIGHT, radius=8, outline=TEAL)
    d.text((36, 175), "Parameters (change these)", font=font(FONT_B, 13), fill=TEAL)
    params = [
        ("Selected Region", "West"),
        ("Selected Category", "Technology"),
        ("Selected Year", "2026"),
        ("High-Value Threshold ($)", "$1,000"),
    ]
    for i, (lab, val) in enumerate(params):
        y = 200 + i * 20
        d.text((36, y), lab, font=font(FONT, 12), fill=DARK)
        d.rectangle((250, y - 2, 400, y + 16), fill=SOFT_AMBER, outline=GOLD)
        d.text((258, y), val, font=font(FONT_B, 12), fill=DARK)

    west_sales = orders.loc[orders["Region"] == "West", "Sales"].sum()
    west_tech = orders.loc[(orders["Region"] == "West") & (orders["Category"] == "Technology"), "Sales"].sum()
    west_2026 = orders.loc[(orders["Region"] == "West") & (orders["Year"] == 2026)].shape[0]
    corp_aov = orders.loc[orders["Segment"] == "Corporate", "Sales"].mean()
    mgr = people.loc[people["Region"] == "West", "Regional Manager"].iloc[0]
    yoy = orders.loc[orders["Year"] == 2026, "Sales"].sum() / orders.loc[orders["Year"] == 2025, "Sales"].sum() - 1
    high_n = orders.loc[(orders["Region"] == "West") & (orders["Sales"] >= 1000)].shape[0]
    peak = orders.loc[orders["Region"] == "West", "Sales"].max()

    rows = [
        ("SUMIFS — Revenue by Region", money(west_sales), "=SUMIFS(Clean_Model!O:O,Clean_Model!F:F,$C$6)"),
        ("SUMIFS — Rev Region+Category", money(west_tech), "=SUMIFS(...F:F,$C$6, ...G:G,$C$7)"),
        ("COUNTIFS — Orders (Region+Year)", f"{west_2026:,}", "=COUNTIFS(...F:F,$C$6, ...C:C,$C$8)"),
        ("AVERAGEIFS — Corporate AOV", money(corp_aov), '=AVERAGEIFS(...O:O,...I:I,"Corporate")'),
        ("XLOOKUP — Manager for Region", mgr, "=XLOOKUP($C$6,Lookups!B:B,Lookups!C:C)"),
        ("YoY — 2026 vs 2025 Revenue", f"{yoy*100:+.1f}%", "=SUMIFS(2026)/SUMIFS(2025)-1"),
        ("COUNTIFS — High value (≥$1k)", f"{high_n:,}", '=COUNTIFS(...O:O,">="&$C$9,...)'),
        ("MAXIFS — Peak order (Region)", money(peak), "=MAXIFS(Clean_Model!O:O,Clean_Model!F:F,$C$6)"),
    ]

    # Table header
    hx, hy = 450, 165
    d.rectangle((hx, hy, W - 20, hy + 28), fill=HEADER)
    d.text((hx + 10, hy + 6), "Function / Pattern", font=font(FONT_B, 11), fill=WHITE)
    d.text((hx + 280, hy + 6), "Result", font=font(FONT_B, 11), fill=WHITE)
    d.text((hx + 400, hy + 6), "Formula (audit)", font=font(FONT_B, 11), fill=WHITE)

    for i, (lab, res, audit) in enumerate(rows):
        y = hy + 28 + i * 36
        bg = SOFT_BLUE if i % 2 == 0 else WHITE
        d.rectangle((hx, y, W - 20, y + 36), fill=bg, outline=GRID)
        d.text((hx + 10, y + 10), lab, font=font(FONT, 11), fill=DARK)
        d.text((hx + 280, y + 10), res, font=font(FONT_B, 12), fill=TEAL)
        d.text((hx + 400, y + 10), audit[:42], font=font(FONT_M, 10), fill=(30, 58, 95))

    d.text(
        (20, H - 56),
        "Also on sheet: INDEX/MATCH, IFS, nested IF bands, SUMPRODUCT, TEXTJOIN, RANK, running totals, FILTER/UNIQUE/SORT (Excel 365).",
        font=font(FONT, 11),
        fill=GRAY,
    )
    out = SHOTS / "advanced_formulas.png"
    img.save(out)
    print("wrote", out)
    return img


def make_lookups(people: pd.DataFrame):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    draw_excel_chrome(d, "Lookups")
    d.text((20, 110), "Lookups — Dimension tables", font=font(FONT_B, 18), fill=NAVY)
    d.text(
        (20, 138),
        "Region→Manager from Superstore People sheet  ·  Category target margins  ·  Segment priority",
        font=font(FONT, 12),
        fill=GRAY,
    )

    # Region manager table
    tables = [
        (
            40,
            "Region → Manager",
            ["Region", "Manager"],
            [[r["Region"], r["Regional Manager"]] for _, r in people.sort_values("Region").iterrows()],
        ),
        (
            420,
            "Category → Target Margin",
            ["Category", "TargetMargin"],
            [
                ["Furniture", "8.0%"],
                ["Office Supplies", "18.0%"],
                ["Technology", "18.0%"],
            ],
        ),
        (
            800,
            "Segment → Priority",
            ["Segment", "Priority"],
            [
                ["Consumer", "P3-Volume"],
                ["Corporate", "P1-Strategic"],
                ["Home Office", "P2-Growth"],
            ],
        ),
    ]
    for x0, title, headers, data in tables:
        d.text((x0, 180), title, font=font(FONT_B, 14), fill=TEAL)
        y = 210
        d.rectangle((x0, y, x0 + 280, y + 28), fill=HEADER)
        d.text((x0 + 8, y + 6), headers[0], font=font(FONT_B, 11), fill=WHITE)
        d.text((x0 + 130, y + 6), headers[1], font=font(FONT_B, 11), fill=WHITE)
        for i, row in enumerate(data):
            yy = y + 28 + i * 32
            bg = LIGHT if i % 2 == 0 else WHITE
            d.rectangle((x0, yy, x0 + 280, yy + 32), fill=bg, outline=GRID)
            d.text((x0 + 8, yy + 8), str(row[0]), font=font(FONT, 12), fill=DARK)
            d.text((x0 + 130, yy + 8), str(row[1]), font=font(FONT, 12), fill=DARK)

    d.text(
        (20, 420),
        "Clean_Model uses XLOOKUP against these ranges for RegionManager, TargetMargin, and Priority on every order line.",
        font=font(FONT, 12),
        fill=GRAY,
    )
    out = SHOTS / "lookups.png"
    img.save(out)
    print("wrote", out)
    return img


def make_clean_model(orders: pd.DataFrame, people: pd.DataFrame):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    draw_excel_chrome(d, "Clean_Model")
    d.text((20, 110), "Clean_Model — formula-driven derived columns", font=font(FONT_B, 18), fill=NAVY)
    d.text(
        (20, 136),
        "Every derived field is an Excel formula referencing Raw_Data + Lookups (Year/Month, Net, Margin, XLOOKUP Manager, flags).",
        font=font(FONT, 12),
        fill=GRAY,
    )

    mgr_map = {r["Region"]: r["Regional Manager"] for _, r in people.iterrows()}
    sample = orders.sort_values("Order Date").head(12).copy()
    sample["Net"] = sample["Sales"]
    sample["Margin"] = sample["Profit"] / sample["Sales"]
    sample["Manager"] = sample["Region"].map(mgr_map)
    sample["HV"] = sample["Sales"].apply(lambda x: "HIGH VALUE" if x >= 1000 else "")
    sample["Loss"] = sample["Profit"].apply(lambda x: "LOSS" if x < 0 else "")

    cols = [
        ("OrderID", 130),
        ("Year", 50),
        ("Region", 70),
        ("Category", 120),
        ("NetRevenue", 95),
        ("Margin%", 75),
        ("Manager (XLOOKUP)", 150),
        ("HighValue", 100),
        ("Loss", 60),
    ]
    x0, y0 = 16, 170
    x = x0
    for name, wcol in cols:
        d.rectangle((x, y0, x + wcol, y0 + 28), fill=HEADER)
        d.text((x + 4, y0 + 6), name, font=font(FONT_B, 10), fill=WHITE)
        x += wcol

    for i, (_, r) in enumerate(sample.iterrows()):
        y = y0 + 28 + i * 30
        bg = LIGHT if i % 2 == 0 else WHITE
        vals = [
            str(r["Order ID"]),
            str(r["Year"]),
            str(r["Region"]),
            str(r["Category"]),
            f"${r['Net']:,.0f}",
            f"{r['Margin']*100:.1f}%",
            str(r["Manager"]),
            str(r["HV"]),
            str(r["Loss"]),
        ]
        x = x0
        for (name, wcol), val in zip(cols, vals):
            fill = bg
            if name == "Margin%":
                if r["Margin"] >= 0.2:
                    fill = (209, 250, 229)
                elif r["Margin"] < 0.1:
                    fill = (254, 226, 226)
            if name == "Loss" and val == "LOSS":
                fill = (254, 226, 226)
            if name == "HighValue" and val:
                fill = SOFT_AMBER
            d.rectangle((x, y, x + wcol, y + 30), fill=fill, outline=GRID)
            d.text((x + 4, y + 8), val, font=font(FONT, 11), fill=DARK)
            x += wcol

    d.text(
        (20, H - 56),
        "Formulas e.g. =XLOOKUP(F2,Lookups!$B$7:$B$10,Lookups!$C$7:$C$10,\"Unknown\")  ·  =IF(O2>=1000,\"HIGH VALUE\",\"\")",
        font=font(FONT_M, 11),
        fill=(30, 58, 95),
    )
    out = SHOTS / "clean_model.png"
    img.save(out)
    print("wrote", out)
    return img


def make_video_frames(dashboard_img, raw_img, adv_img):
    FRAMES.mkdir(parents=True, exist_ok=True)
    # Title
    img = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(img)
    d.rectangle((0, H - 120, W, H), fill=TEAL)
    d.text((64, 180), "Sales Advanced Excel", font=font(FONT_B, 48), fill=WHITE)
    d.text((64, 250), "Portfolio Dashboard", font=font(FONT_B, 36), fill=GOLD)
    d.text(
        (64, 330),
        "Lookups · Clean Model · Advanced Formulas · Executive KPIs",
        font=font(FONT, 20),
        fill=(200, 220, 220),
    )
    d.text(
        (64, H - 78),
        "GitHub: shanitripathi417-code  ·  Tableau Sample Superstore  ·  USD",
        font=font(FONT, 18),
        fill=WHITE,
    )
    img.save(FRAMES / "01_title.png")

    raw_img.save(FRAMES / "02_data.png")
    adv_img.save(FRAMES / "03_formulas.png")
    dashboard_img.save(FRAMES / "04_dashboard.png")

    # End card
    img = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(img)
    d.rectangle((0, H - 120, W, H), fill=TEAL)
    d.text((64, 200), "Thank you", font=font(FONT_B, 52), fill=WHITE)
    d.text((64, 280), "Open the .xlsx in Excel to explore live formulas", font=font(FONT, 22), fill=GOLD)
    d.line((64, 320, 640, 320), fill=GOLD, width=2)
    d.text(
        (64, H - 78),
        "Sales_Advanced_Excel_Dashboard.xlsx  ·  shanitripathi417-code",
        font=font(FONT, 18),
        fill=WHITE,
    )
    img.save(FRAMES / "05_end.png")
    print("wrote video frames")


def remux_video():
    import subprocess

    scenes = [
        (FRAMES / "01_title.png", 4.0),
        (FRAMES / "02_data.png", 5.0),
        (FRAMES / "03_formulas.png", 6.0),
        (FRAMES / "04_dashboard.png", 6.0),
        (FRAMES / "05_end.png", 4.0),
    ]
    out = ROOT / "artifacts" / "sales-advanced-excel-demo.mp4"
    inputs = []
    for path, dur in scenes:
        inputs.extend(["-loop", "1", "-t", f"{dur:.2f}", "-i", str(path)])
    n = len(scenes)
    filt = "".join(f"[{i}:v]" for i in range(n)) + f"concat=n={n}:v=1:a=0[v]"
    cmd = [
        "ffmpeg",
        "-y",
        *inputs,
        "-filter_complex",
        filt,
        "-map",
        "[v]",
        "-r",
        "30",
        "-c:v",
        "libx264",
        "-pix_fmt",
        "yuv420p",
        "-preset",
        "medium",
        "-crf",
        "23",
        "-movflags",
        "+faststart",
        str(out),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stderr[-2000:])
        raise SystemExit(r.returncode)
    print("wrote", out, "bytes", out.stat().st_size)


def main():
    SHOTS.mkdir(parents=True, exist_ok=True)
    orders, people = load()
    dash = make_dashboard(orders)
    raw = make_raw_preview(orders)
    adv = make_advanced(orders, people)
    make_lookups(people)
    make_clean_model(orders, people)
    make_video_frames(dash, raw, adv)
    remux_video()
    print("done")


if __name__ == "__main__":
    main()
