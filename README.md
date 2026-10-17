# Retail Sales & Profitability Executive Dashboard (Microsoft Excel 2021)

An end-to-end commercial retail business intelligence workbook engineered in Microsoft Excel 2021, featuring dynamic financial KPI cards, cross-sheet mathematical modeling (`SUMIFS`, `COUNTIFS`, `XLOOKUP`, `IF`), category and regional drill-downs, and automated native charts.

---

## 1. Project Overview & Business Problem
Retail enterprise leadership requires real-time visibility into sales volume, net profitability, discounting efficiency, and regional performance. Without consolidated analytical dashboards, executives struggle to identify loss-making merchandise categories, over-discounted product lines, or territories suffering from fulfillment overhead.

This project delivers a multi-tab financial dashboard workbook modeled on 9,994 retail transactions across the United States. Designed specifically for **Microsoft Excel 2021**, the workbook eliminates manual calculation by using dynamic array functions, structured formula models, and native chart visualizations.

---

## 2. Business Objectives & Analytical Questions
1. **Financial Health**: What is the company's gross revenue, net profit, and overall profit margin?
2. **Category Profitability**: Which product categories generate strong revenue but deliver negative margins due to excessive discounting?
3. **Regional Distribution**: How does sales contribution divide across the West, East, Central, and South regions?
4. **Time-Series Momentum**: How has monthly revenue trended across the 4-year operating timeline?
5. **Product Sub-Category Triage**: How are individual sub-categories classified using automated conditional flags (`Profitable` vs. `Loss-Making`)?

---

## 3. Technology & Compatibility
- **Application**: Microsoft Excel 2021 (Home & Student / Professional / Microsoft 365)
- **Engine**: Native Excel calculation engine, OpenPyXL 3.1.5 (automated workbook build)
- **Advanced Formula Functions**: `SUMIFS`, `COUNTIFS`, `XLOOKUP`, `IF`, `YEAR`, `TEXT`, dynamic array division.

---

## 4. Repository & Workbook Architecture
```text
retail-sales-excel-dashboard/
├── data/
│   └── superstore.csv                  # Raw retail transactions dataset (9,994 records)
├── create_retail_workbook.py           # Automated openpyxl generation script
├── retail_sales_dashboard.xlsx         # Production Excel 2021 workbook
├── PROJECT_WALKTHROUGH.md              # Formula breakdown & Excel exercise guide
└── README.md                           # Complete documentation & interview guide
```

### Workbook Sheets Structure
| Sheet Tab Name | Purpose & Contents |
| :--- | :--- |
| **`Executive Dashboard`** | High-level KPI cards (Total Sales, Total Profit, Margin %, Order Count, Units Sold), Category Summary Table & Column Chart, Regional Summary Table & Doughnut Chart. |
| **`Monthly Trends`** | Time-series monthly table (48 months) tracking Revenue, Profit, Margin %, and Order Count via `SUMIFS` with a native Line Chart. |
| **`Sub-Category Breakdown`** | Granular 17-row performance table utilizing `XLOOKUP` for parent category mapping, `SUMIFS` for revenue/profit, and `IF()` for profitability status flags. |
| **`Orders Data`** | Master operational dataset of 9,994 formatted transaction rows with formulated columns (`Cost = Sales - Profit`, `Profit Margin = Profit / Sales`). |
| **`Data Quality & Dictionary`** | Full metadata definitions, data types, formula logic, and data boundary validations. |

---

## 5. Dataset Attribution & Core Metrics

- **Dataset Source**: Public Benchmark Superstore Retail Dataset
- **Geographic Scope**: United States Commercial Operations
- **Total Orders Recorded**: 9,994 transactions

### Verified Empirical Findings (Live Excel Calculations)
```text
Total Gross Revenue:     $2,297,200.86
Total Net Profit:        $286,397.02
Overall Profit Margin:   12.47%
Total Units Sold:        37,873 units
Total Transactions:      9,994 lines (5,009 unique orders)
```

### Category Financial Summary
| Category | Gross Sales ($) | Net Profit ($) | Profit Margin (%) | Units Sold |
| :--- | :--- | :--- | :--- | :--- |
| **Technology** | $836,154.03 | $145,454.95 | **17.39%** | 6,939 |
| **Office Supplies** | $719,047.03 | $122,490.80 | **17.03%** | 22,906 |
| **Furniture** | $741,999.80 | $18,451.27 | **2.49%** | 8,028 |

> [!WARNING]
> **Furniture Margin Warning**: While Furniture generates nearly identical gross sales ($742k) to Office Supplies ($719k), its net profit ($18.5k) is severely depressed (2.49% margin) primarily due to severe discounting and high freight costs on Tables and Bookcases.

---

## 6. Formula Architecture & Implementation Guide

### 1. Multi-Condition Aggregation (`SUMIFS`)
```excel
=SUMIFS('Orders Data'!R:R, 'Orders Data'!O:O, B11)
```
- Sums the `Sales` column (`R:R`) where the `Category` column (`O:O`) equals the cell reference `B11` ("Furniture").

### 2. Modern Lookups (`XLOOKUP`)
```excel
=XLOOKUP(B5, 'Orders Data'!P:P, 'Orders Data'!O:O, "Unknown")
```
- Looks up the Sub-Category in `B5` (e.g. "Chairs") within column `P:P` and returns its parent category from column `O:O` without column count indexing errors common to legacy `VLOOKUP`.

### 3. Conditional Business Logic (`IF`)
```excel
=IF(E5 > 0, "Profitable", "Loss-Making")
```
- Evaluates whether net profit for a sub-category is positive. Identifies loss-making lines (such as Tables at -$17,725.48 net loss) automatically.

---

## 7. Manual Verification & Excel 2021 Usage
1. Open `retail_sales_dashboard.xlsx` in **Microsoft Excel 2021**.
2. Click **Enable Editing** if opened in Protected View.
3. Verify that KPI card values in `Executive Dashboard!B6:K6` display `$2,297,200.86`, `$286,397.02`, and `12.5%`.
4. Review the two embedded charts:
   - Category Bar Chart (`B16:F30`).
   - Regional Doughnut Chart (`H16:L30`).
5. Navigate to the `Monthly Trends` sheet to inspect the chronological revenue progression.

---

## 8. Interview Preparation: 10 Core Excel Questions & Answers

### Q1: Why is `XLOOKUP` superior to `VLOOKUP` in Excel 2021?
**Answer**:
1. `VLOOKUP` requires the lookup column to be the first column on the left; `XLOOKUP` can look in any direction (left or right).
2. `VLOOKUP` uses static column index numbers (e.g., column 5), which break if new columns are inserted; `XLOOKUP` uses dynamic range references.
3. `XLOOKUP` defaults to exact match (`match_mode=0`), eliminating the common `#N/A` error caused by omitting `FALSE` in `VLOOKUP`.

### Q2: How does `SUMIFS` differ from `SUMIF`?
**Answer**: `SUMIF` only supports a single condition and has its sum range as the last argument: `SUMIF(criteria_range, criteria, sum_range)`. `SUMIFS` supports up to 127 criteria pairs and places the `sum_range` as the first argument: `SUMIFS(sum_range, criteria_range1, criteria1, ...)`, establishing consistent syntax across `SUMIFS`, `AVERAGEIFS`, and `COUNTIFS`.

### Q3: What is the risk of using full-column references like `'Orders Data'!R:R`?
**Answer**: In older Excel versions, referencing an entire column forced Excel to evaluate 1,048,576 rows, causing workbook lag. In Excel 2021 and modern versions, Excel's calculation engine optimizes full-column references to evaluate only the used range (`UsedRange`). However, inside structured tables, referencing named table columns (e.g. `Table1[Sales]`) is preferred for performance and self-documenting readability.

### Q4: Why did Furniture generate such a low profit margin (2.49%) compared to Technology (17.39%)?
**Answer**: Our sub-category breakdown reveals that Tables generated a net loss of over -$17,700 and Bookcases yielded marginal returns due to promotional discounts exceeding 40% on bulky merchandise. Technology items (Phones, Copiers) maintain higher margins and lower price discounts.

### Q5: What is the purpose of the number format `$#,##0.00;($#,##0.00);"-"`?
**Answer**: This standard financial accounting format renders positive numbers as `$1,250.00`, negative numbers in parenthetical accounting syntax as `($1,250.00)`, and zero values as a clean hyphen `-`, preventing visual clutter in financial statements.

### Q6: How do you handle `#DIV/0!` errors when calculating profit margins in Excel?
**Answer**: By wrapping the division in `IFERROR()`:
```excel
=IFERROR(Profit / Sales, 0)
```
or by checking if sales are positive using `IF()`:
```excel
=IF(Sales > 0, Profit / Sales, 0)
```

### Q7: What are Excel structured table references (ListObjects)?
**Answer**: Structured references replace coordinate cell notation (like `A2:A1000`) with table and column names (like `SalesData[Sales]`). When new rows are appended to the table, formulas and pivot tables referencing structured tables automatically expand without needing manual range adjustments.

### Q8: How did you ensure this workbook is fully compatible with Microsoft Excel 2021?
**Answer**: We restricted all formulas to functions natively supported in the permanent perpetual release of Excel 2021 (`SUMIFS`, `COUNTIFS`, `XLOOKUP`, `IF`, `TEXT`), avoiding Microsoft 365 insider-only functions (like `GROUPBY` or `PIVOTBY`).

### Q9: What is the difference between a Pivot Table and a formulated summary table?
**Answer**: Pivot Tables provide rapid, drag-and-drop interactive aggregation and built-in slicing, but require manual cache refresh (`Alt + F5`) when data changes. Formulated summary tables (`SUMIFS`, `XLOOKUP`) update instantaneously and allow custom cell-by-cell typography, layout formatting, and embedded KPI card styling.

### Q10: How would you add interactivity to this dashboard for executive presentations?
**Answer**: I would insert Excel Slicers connected to Region, Segment, and Order Year, allowing the user to filter the dashboard with a single click while preserving the structured layout.
