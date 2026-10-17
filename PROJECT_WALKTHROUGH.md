# Project Walkthrough: Retail Sales Dashboard in Microsoft Excel 2021

## 1. High-Level Workbook Architecture
```
Master Operational Data ('Orders Data', 9,994 Rows)
  ├── Date Parsing & Formulated Metrics (Cost, Profit Margin, YearMonth)
  ├── Sub-Category Aggregation via XLOOKUP, SUMIFS, and IF
  ├── Chronological Monthly Time-Series ('Monthly Trends')
  └── Executive Visual Dashboard
       ├── 5 Formulated KPI Cards ($2.3M Sales, $286k Profit, 12.5% Margin)
       ├── Category Performance Column Chart
       └── Regional Sales Distribution Doughnut Chart
```

---

## 2. Core Excel Formulas Explained

### Modern XLOOKUP for Relational Mapping
```excel
=XLOOKUP(B5, 'Orders Data'!P:P, 'Orders Data'!O:O, "Unknown")
```
- Looks up the sub-category in cell `B5` (e.g., "Chairs") against column `P:P` of `Orders Data`.
- Retrieves the corresponding major Category from column `O:O` ("Furniture").
- Returns `"Unknown"` if no match is found, eliminating `#N/A` errors.

### Conditional Profitability Flag
```excel
=IF(E5 > 0, "Profitable", "Loss-Making")
```
- Evaluates whether net profit in cell `E5` exceeds zero.
- Flags problematic lines (such as Tables, which generated -$17,725.48 net loss) for managerial review.

---

## 3. Three Improvements You Can Personally Implement

1. **Insert Excel Slicers for Interactive Filtering**:
   - In Excel 2021, select the `Orders Data` table -> **Table Design** -> **Insert Slicer**.
   - Select **Region** and **Segment** to enable instant 1-click filtering of the dashboard.
2. **Conditional Formatting Data Bars & Icons**:
   - On the `Sub-Category Breakdown` sheet, highlight the `Margin %` column -> **Conditional Formatting** -> **Data Bars** (Green for positive, Red for negative).
3. **What-If Parameter Table for Discount Reduction**:
   - Add a scenario sensitivity model showing how reducing average discount by 5% on Furniture would boost overall company profit by over $35,000.

---

## 4. Hands-On Practice Exercises
- **Exercise 1**: Write a formula on the dashboard using `AVERAGEIFS` that computes the average order discount specifically for the "Corporate" customer segment in the "West" region.
- **Exercise 2**: Use the Excel `INDEX / MATCH` combination instead of `XLOOKUP` to verify backwards compatibility with Excel 2016.
- **Exercise 3**: Create a Sparkline in cell `G11` of the Executive Dashboard showing the 4-year trend for the Furniture category.
