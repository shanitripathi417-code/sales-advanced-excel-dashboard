# Sales Advanced Excel — Portfolio Project

**GitHub:** [shanitripathi417-code](https://github.com/shanitripathi417-code)

Multi-sheet Excel workbook for retail sales analysis: fact/dimension model, advanced formulas, conditional formatting, and an executive dashboard. KPIs and chart inputs are driven by **live Excel formulas** (not static pasted totals).

| Artifact | Path |
|---|---|
| Workbook | [`Sales_Advanced_Excel_Dashboard.xlsx`](./Sales_Advanced_Excel_Dashboard.xlsx) |
| Demo video (silent) | [`artifacts/sales-advanced-excel-demo.mp4`](./artifacts/sales-advanced-excel-demo.mp4) |
| Screenshots | [`screenshots/`](./screenshots/) |
| Source data | [`data/sample_superstore_orders.csv`](./data/sample_superstore_orders.csv) |

> Open in **Microsoft Excel** (desktop or Excel 365). Formulas recalculate when you open the file. Currency: **USD**.

---

## What this project covers

- Fact / dimension layout: `Raw_Data` + `Lookups` → `Clean_Model` → `Dashboard`
- Lookups: **XLOOKUP**, **VLOOKUP**, **INDEX/MATCH**
- Aggregations: **SUMIFS / COUNTIFS / AVERAGEIFS / MAXIFS / MINIFS**
- Logic: **IF / IFS**, high-value and discount bands
- Dates: **YEAR / MONTH / TEXT / EOMONTH / DATE**
- Other patterns: **SUMPRODUCT**, **TEXTJOIN**, **RANK**, running totals, YoY
- Excel 365 (optional): **UNIQUE / SORT / FILTER**
- Conditional formatting on margin and margin vs target
- Dashboard charts: Region, Category, Monthly trend, Segment mix

---

## Sheet map

| # | Sheet | Purpose |
|---|---|---|
| 1 | **Instructions** | How to use the workbook + data attribution |
| 2 | **Dashboard** | KPI cards + 4 charts |
| 3 | **Raw_Data** | **10,194** Sample Superstore order lines (2023–2026) |
| 4 | **Lookups** | Region→Manager, Category→Target Margin, Segment→Priority |
| 5 | **Clean_Model** | Derived columns via formulas on `Raw_Data` + `Lookups` |
| 6 | **Advanced_Formulas** | Parameter panel with working formula demos |

```
Raw_Data (facts) ──► Clean_Model (formulas + XLOOKUP) ──► Dashboard (KPIs/charts)
Lookups (dims)  ──┘                          └──► Advanced_Formulas
```

---

## Formulas used

### `Clean_Model`

| Pattern | Role |
|---|---|
| `YEAR` / `MONTH` / `TEXT` | Calendar fields from `OrderDate` |
| Arithmetic | Gross, Net after discount, Profit |
| `IF` | Margin %; high-value flag (≥ **$1,000**) |
| `XLOOKUP` | Manager, target margin, segment priority |
| `IFS` | Discount band |
| Conditional formatting | Margin thresholds; margin vs target |

### `Advanced_Formulas`

Amber inputs: **C6** Region · **C7** Category · **C8** Year · **C9** high-value threshold.

| Function | Use |
|---|---|
| **SUMIFS** | Revenue by Region; Region+Category |
| **COUNTIFS** | Orders by Region+Year; high-value count |
| **AVERAGEIFS** | Segment AOV; margin by Category+Year |
| **MAXIFS / MINIFS** | Largest / smallest order |
| **XLOOKUP** / **VLOOKUP** | Region manager |
| **INDEX / MATCH** | Category target margin |
| **IFS** / nested **IF** | Tier / performance band |
| **TEXT** / **DATE** / **EOMONTH** | Period labels and month-end |
| **SUMPRODUCT** | Unique-customer style count; weighted margin |
| **TEXTJOIN** | Region list |
| **RANK** | Regions by revenue |
| Running total / **YoY** | Cumulative month revenue; year-over-year via SUMIFS |
| **IFERROR** | Safe divides |

Excel 365 dynamic arrays (`UNIQUE`, `SORT`, `FILTER`) are included at the bottom of that sheet. Older Excel may show `#NAME?` for those only.

---

## Screenshots

### Dashboard

![Dashboard](screenshots/dashboard.png)

### Advanced formulas

![Advanced Formulas](screenshots/advanced_formulas.png)

### Raw data

![Raw Data](screenshots/raw_data_preview.png)

### Lookups

![Lookups](screenshots/lookups.png)

### Clean model

![Clean Model](screenshots/clean_model.png)

---

## Data source

| | |
|---|---|
| **Dataset** | [Tableau Sample Superstore](https://public.tableau.com/app/learn/sample-data) (Orders + People) |
| **Download** | [sample_-_superstore.xls](https://public.tableau.com/app/sample-data/sample_-_superstore.xls) |
| **Local files** | `data/sample_superstore.xls`, `data/sample_superstore_orders.csv`, `data/sample_superstore_people.csv` |
| **Rows** | **10,194** order lines |
| **Dates** | 2023-01-03 → 2026-12-30 |
| **Currency** | USD |
| **Regions** | Central, East, South, West |
| **Categories** | Furniture, Office Supplies, Technology |
| **Segments** | Consumer, Corporate, Home Office |

### Field mapping

| Superstore | Workbook |
|---|---|
| Order ID | OrderID |
| Order Date | OrderDate |
| Region | Region |
| Category | ProductCategory |
| Product Name | Product |
| Segment | CustomerSegment |
| Customer Name | Customer |
| Quantity | Quantity |
| Discount | Discount |
| Sales | Revenue |
| Profit | Profit |
| State/Province | State |
| Sub-Category | SubCategory |

Target margins on `Lookups` (Furniture 8%, Office Supplies / Technology 18%) are workbook assumptions for margin-vs-target demos, not fields from Tableau.

---

## How to open

1. Download [`Sales_Advanced_Excel_Dashboard.xlsx`](./Sales_Advanced_Excel_Dashboard.xlsx)
2. Open in Microsoft Excel
3. Start on **Dashboard**, then try amber parameters on **Advanced_Formulas**
4. Optional: [`artifacts/sales-advanced-excel-demo.mp4`](./artifacts/sales-advanced-excel-demo.mp4)

---

## License

Portfolio project for [shanitripathi417-code](https://github.com/shanitripathi417-code). Retail rows from Tableau’s public Sample Superstore. Not an employer confidential dataset.
