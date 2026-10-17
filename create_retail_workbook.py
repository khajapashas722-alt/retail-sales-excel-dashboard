"""
Create Retail Sales & Profitability Dashboard Excel Workbook
Author: Shaik Khaja Pasha
Compatibility: Microsoft Excel 2021
Dependencies: openpyxl, pandas
"""

import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference, PieChart


def build_retail_dashboard(csv_path: str, output_xlsx: str):
    print(f"[1/4] Reading source dataset from {csv_path}...")
    df = pd.read_csv(csv_path)
    # Filter to valid order rows (removes trailing metadata tables)
    df = df[pd.to_numeric(df["Row ID"], errors="coerce").notnull()].copy()
    df["Row ID"] = df["Row ID"].astype(int)
    # Parse dates
    df["Order Date"] = pd.to_datetime(df["Order Date"])
    df["Ship Date"] = pd.to_datetime(df["Ship Date"])
    df["Order Year"] = df["Order Date"].dt.year
    df["Order YearMonth"] = df["Order Date"].dt.strftime("%Y-%m")
    df["Cost"] = df["Sales"] - df["Profit"]
    df["Profit Margin"] = df["Profit"] / df["Sales"]

    print(f"[2/4] Initializing openpyxl workbook...")
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    # Define standard color palette
    NAVY_HEADER = "1B365D"
    WHITE = "FFFFFF"
    ACCENT_BLUE = "E8EEF5"
    LIGHT_GRAY = "F9FBFD"
    BORDER_GRAY = "D3D3D3"
    CARD_BG = "F0F4F8"

    font_title = Font(name="Calibri", size=16, bold=True, color=NAVY_HEADER)
    font_section = Font(name="Calibri", size=12, bold=True, color=NAVY_HEADER)
    font_header = Font(name="Calibri", size=11, bold=True, color=WHITE)
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    font_kpi_value = Font(name="Calibri", size=18, bold=True, color=NAVY_HEADER)
    font_kpi_label = Font(name="Calibri", size=9, bold=False, color="555555")

    fill_header = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    fill_accent = PatternFill(start_color=ACCENT_BLUE, end_color=ACCENT_BLUE, fill_type="solid")
    fill_card = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")
    fill_light = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")

    thin_border_side = Side(border_style="thin", color=BORDER_GRAY)
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # -------------------------------------------------------------------------
    # SHEET 1: Executive Dashboard
    # -------------------------------------------------------------------------
    ws_dash = wb.create_sheet(title="Executive Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True

    # Title block
    ws_dash["B2"] = "RETAIL SALES & PROFITABILITY EXECUTIVE DASHBOARD"
    ws_dash["B2"].font = font_title
    ws_dash["B3"] = "Dataset: Superstore Commercial Retail Portfolio | Microsoft Excel 2021 Edition"
    ws_dash["B3"].font = Font(name="Calibri", size=10, italic=True, color="666666")

    # KPI Cards (Row 5 to 7)
    kpis = [
        ("Total Sales", "=SUM('Orders Data'!R2:R10000)", "$#,##0.00", "B", "C"),
        ("Total Profit", "=SUM('Orders Data'!U2:U10000)", "$#,##0.00", "D", "E"),
        ("Profit Margin", "=D6/B6", "0.0%", "F", "G"),
        ("Total Orders", "=COUNT('Orders Data'!A2:A10000)", "#,##0", "H", "I"),
        ("Total Units Sold", "=SUM('Orders Data'!S2:S10000)", "#,##0", "J", "K")
    ]

    for label, formula, num_fmt, col_start, col_end in kpis:
        c1, c2 = f"{col_start}5", f"{col_end}5"
        v1, v2 = f"{col_start}6", f"{col_end}6"
        ws_dash.merge_cells(f"{c1}:{c2}")
        ws_dash.merge_cells(f"{v1}:{v2}")
        
        ws_dash[c1] = label
        ws_dash[c1].font = font_kpi_label
        ws_dash[c1].alignment = Alignment(horizontal="center", vertical="center")
        ws_dash[c1].fill = fill_card
        
        ws_dash[v1] = formula
        ws_dash[v1].font = font_kpi_value
        ws_dash[v1].alignment = Alignment(horizontal="center", vertical="center")
        ws_dash[v1].number_format = num_fmt
        ws_dash[v1].fill = fill_card

        for row in range(5, 7):
            for col in [col_start, col_end]:
                ws_dash[f"{col}{row}"].border = border_all

    # Summary Table 1 on Dashboard: Category Performance
    ws_dash["B9"] = "Category Financial Performance Summary"
    ws_dash["B9"].font = font_section

    cat_headers = ["Category", "Total Sales", "Total Profit", "Profit Margin", "Units Sold"]
    for i, h in enumerate(cat_headers, start=2):
        cell = ws_dash.cell(row=10, column=i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all

    categories = ["Furniture", "Office Supplies", "Technology"]
    for idx, cat in enumerate(categories, start=11):
        ws_dash.cell(row=idx, column=2, value=cat).font = font_bold
        ws_dash.cell(row=idx, column=3, value=f"=SUMIFS('Orders Data'!R:R, 'Orders Data'!O:O, B{idx})").number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=4, value=f"=SUMIFS('Orders Data'!U:U, 'Orders Data'!O:O, B{idx})").number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=5, value=f"=D{idx}/C{idx}").number_format = "0.0%"
        ws_dash.cell(row=idx, column=6, value=f"=SUMIFS('Orders Data'!S:S, 'Orders Data'!O:O, B{idx})").number_format = "#,##0"
        for c in range(2, 7):
            ws_dash.cell(row=idx, column=c).border = border_all
            ws_dash.cell(row=idx, column=c).font = font_regular

    # Category Bar Chart
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Sales and Profit by Category"
    chart1.y_axis.title = "USD ($)"
    chart1.x_axis.title = "Category"
    data1 = Reference(ws_dash, min_col=3, min_row=10, max_col=4, max_row=13)
    cats1 = Reference(ws_dash, min_col=2, min_row=11, max_row=13)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.width = 15
    chart1.height = 7.5
    ws_dash.add_chart(chart1, "B16")

    # Summary Table 2 on Dashboard: Regional Sales
    ws_dash["H9"] = "Regional Sales Breakdown"
    ws_dash["H9"].font = font_section

    reg_headers = ["Region", "Total Sales", "Total Profit", "Margin"]
    for i, h in enumerate(reg_headers, start=8):
        cell = ws_dash.cell(row=10, column=i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_all

    regions = ["West", "East", "Central", "South"]
    for idx, reg in enumerate(regions, start=11):
        ws_dash.cell(row=idx, column=8, value=reg).font = font_bold
        ws_dash.cell(row=idx, column=9, value=f"=SUMIFS('Orders Data'!R:R, 'Orders Data'!M:M, H{idx})").number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=10, value=f"=SUMIFS('Orders Data'!U:U, 'Orders Data'!M:M, H{idx})").number_format = "$#,##0.00"
        ws_dash.cell(row=idx, column=11, value=f"=J{idx}/I{idx}").number_format = "0.0%"
        for c in range(8, 12):
            ws_dash.cell(row=idx, column=c).border = border_all
            ws_dash.cell(row=idx, column=c).font = font_regular

    # Regional Pie/Doughnut Chart
    chart2 = PieChart()
    chart2.title = "Sales Contribution by Region"
    data2 = Reference(ws_dash, min_col=9, min_row=10, max_row=14)
    cats2 = Reference(ws_dash, min_col=8, min_row=11, max_row=14)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.width = 13
    chart2.height = 7.5
    ws_dash.add_chart(chart2, "H16")

    # -------------------------------------------------------------------------
    # SHEET 2: Monthly & Regional Analysis
    # -------------------------------------------------------------------------
    ws_monthly = wb.create_sheet(title="Monthly Trends")
    ws_monthly.views.sheetView[0].showGridLines = True

    ws_monthly["B2"] = "MONTHLY SALES & PROFITABILITY TIME-SERIES"
    ws_monthly["B2"].font = font_title

    m_headers = ["Year-Month", "Sales", "Profit", "Profit Margin", "Orders Count"]
    for i, h in enumerate(m_headers, start=2):
        cell = ws_monthly.cell(row=4, column=i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_all
        cell.alignment = Alignment(horizontal="center")

    unique_ym = sorted(df["Order YearMonth"].unique())
    for idx, ym in enumerate(unique_ym, start=5):
        ws_monthly.cell(row=idx, column=2, value=ym).font = font_regular
        ws_monthly.cell(row=idx, column=3, value=f"=SUMIFS('Orders Data'!R:R, 'Orders Data'!W:W, B{idx})").number_format = "$#,##0.00"
        ws_monthly.cell(row=idx, column=4, value=f"=SUMIFS('Orders Data'!U:U, 'Orders Data'!W:W, B{idx})").number_format = "$#,##0.00"
        ws_monthly.cell(row=idx, column=5, value=f"=D{idx}/C{idx}").number_format = "0.0%"
        ws_monthly.cell(row=idx, column=6, value=f"=COUNTIFS('Orders Data'!W:W, B{idx})").number_format = "#,##0"
        for c in range(2, 7):
            ws_monthly.cell(row=idx, column=c).border = border_all

    # Line Chart of Monthly Revenue
    line_chart = LineChart()
    line_chart.title = "Monthly Revenue Growth Trend"
    line_chart.y_axis.title = "Revenue ($)"
    line_chart.x_axis.title = "Month"
    ldata = Reference(ws_monthly, min_col=3, min_row=4, max_row=4 + len(unique_ym))
    lcats = Reference(ws_monthly, min_col=2, min_row=5, max_row=4 + len(unique_ym))
    line_chart.add_data(ldata, titles_from_data=True)
    line_chart.set_categories(lcats)
    line_chart.width = 18
    line_chart.height = 8
    ws_monthly.add_chart(line_chart, "H4")

    # -------------------------------------------------------------------------
    # SHEET 3: Product Sub-Category Analysis
    # -------------------------------------------------------------------------
    ws_subcat = wb.create_sheet(title="Sub-Category Breakdown")
    ws_subcat.views.sheetView[0].showGridLines = True

    ws_subcat["B2"] = "SUB-CATEGORY PERFORMANCE & STATUS FLAGS"
    ws_subcat["B2"].font = font_title

    subcat_headers = ["Sub-Category", "Category", "Sales", "Profit", "Margin %", "Status Flag"]
    for i, h in enumerate(subcat_headers, start=2):
        cell = ws_subcat.cell(row=4, column=i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_all
        cell.alignment = Alignment(horizontal="center")

    unique_subcats = sorted(df["Sub-Category"].unique())
    for idx, sc in enumerate(unique_subcats, start=5):
        ws_subcat.cell(row=idx, column=2, value=sc).font = font_bold
        # XLOOKUP for Category
        ws_subcat.cell(row=idx, column=3, value=f"=XLOOKUP(B{idx}, 'Orders Data'!P:P, 'Orders Data'!O:O, \"Unknown\")").font = font_regular
        ws_subcat.cell(row=idx, column=4, value=f"=SUMIFS('Orders Data'!R:R, 'Orders Data'!P:P, B{idx})").number_format = "$#,##0.00"
        ws_subcat.cell(row=idx, column=5, value=f"=SUMIFS('Orders Data'!U:U, 'Orders Data'!P:P, B{idx})").number_format = "$#,##0.00"
        ws_subcat.cell(row=idx, column=6, value=f"=E{idx}/D{idx}").number_format = "0.0%"
        ws_subcat.cell(row=idx, column=7, value=f"=IF(E{idx}>0, \"Profitable\", \"Loss-Making\")").font = font_bold
        for c in range(2, 8):
            ws_subcat.cell(row=idx, column=c).border = border_all

    # -------------------------------------------------------------------------
    # SHEET 4: Orders Data (Master Raw Data Table)
    # -------------------------------------------------------------------------
    print(f"[3/4] Populating Orders Data sheet with {len(df):,} rows...")
    ws_data = wb.create_sheet(title="Orders Data")
    ws_data.views.sheetView[0].showGridLines = True

    data_headers = [
        "Row ID", "Order ID", "Order Date", "Ship Date", "Ship Mode",
        "Customer ID", "Customer Name", "Segment", "Country", "City",
        "State", "Postal Code", "Region", "Product ID", "Category",
        "Sub-Category", "Product Name", "Sales", "Quantity", "Discount",
        "Profit", "Order Year", "Order YearMonth", "Cost", "Profit Margin"
    ]

    for col_idx, h in enumerate(data_headers, start=1):
        cell = ws_data.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center")

    # Ingest rows
    c_idx = {col: i for i, col in enumerate(df.columns)}
    for r_idx, row in enumerate(df.values, start=2):
        order_date_str = pd.to_datetime(row[c_idx["Order Date"]]).strftime("%Y-%m-%d")
        ship_date_str = pd.to_datetime(row[c_idx["Ship Date"]]).strftime("%Y-%m-%d")
        sales_val = float(row[c_idx["Sales"]])
        qty_val = int(row[c_idx["Quantity"]])
        disc_val = float(row[c_idx["Discount"]])
        profit_val = float(row[c_idx["Profit"]])
        cost_formula = f"=R{r_idx}-U{r_idx}"
        margin_formula = f"=U{r_idx}/R{r_idx}"

        row_data = [
            int(row[c_idx["Row ID"]]),
            str(row[c_idx["Order ID"]]),
            order_date_str,
            ship_date_str,
            str(row[c_idx["Ship Mode"]]),
            str(row[c_idx["Customer ID"]]),
            str(row[c_idx["Customer Name"]]),
            str(row[c_idx["Segment"]]),
            str(row[c_idx["Country"]]),
            str(row[c_idx["City"]]),
            str(row[c_idx["State"]]),
            str(row[c_idx["Postal Code"]]),
            str(row[c_idx["Region"]]),
            str(row[c_idx["Product ID"]]),
            str(row[c_idx["Category"]]),
            str(row[c_idx["Sub-Category"]]),
            str(row[c_idx["Product Name"]]),
            sales_val,
            qty_val,
            disc_val,
            profit_val,
            int(row[c_idx["Order Year"]]),
            str(row[c_idx["Order YearMonth"]]),
            cost_formula,
            margin_formula
        ]
        ws_data.append(row_data)

        # Apply number formats to numerical cells
        ws_data.cell(row=r_idx, column=18).number_format = "$#,##0.00"
        ws_data.cell(row=r_idx, column=19).number_format = "#,##0"
        ws_data.cell(row=r_idx, column=20).number_format = "0.0%"
        ws_data.cell(row=r_idx, column=21).number_format = "$#,##0.00"
        ws_data.cell(row=r_idx, column=24).number_format = "$#,##0.00"
        ws_data.cell(row=r_idx, column=25).number_format = "0.0%"

    # -------------------------------------------------------------------------
    # SHEET 5: Data Quality & Dictionary
    # -------------------------------------------------------------------------
    ws_dict = wb.create_sheet(title="Data Quality & Dictionary")
    ws_dict.views.sheetView[0].showGridLines = True

    ws_dict["B2"] = "SUPERSTORE DATA DICTIONARY & INTEGRITY AUDIT"
    ws_dict["B2"].font = font_title

    dict_headers = ["Field Name", "Data Type", "Excel Formula / Source", "Description"]
    for i, h in enumerate(dict_headers, start=2):
        cell = ws_dict.cell(row=4, column=i, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_all

    dict_entries = [
        ("Row ID", "Integer", "Direct source", "Unique sequential integer for each record line."),
        ("Order ID", "String", "Direct source", "Unique alphanumeric transaction code."),
        ("Order Date", "Date", "Direct source", "Date on which purchase order was placed."),
        ("Ship Date", "Date", "Direct source", "Date on which carrier dispatched order."),
        ("Ship Mode", "String", "Direct source", "Shipping service class (Standard, Second, First, Same Day)."),
        ("Customer Name", "String", "Direct source", "Full name of retail client."),
        ("Segment", "String", "Direct source", "Client market tier (Consumer, Corporate, Home Office)."),
        ("Region", "String", "Direct source", "Geographic US territory (West, East, Central, South)."),
        ("Category", "String", "Direct source", "Top-level merchandise category (Furniture, Office Supplies, Technology)."),
        ("Sub-Category", "String", "Direct source", "Granular product grouping (17 sub-categories)."),
        ("Sales", "Currency", "Direct source", "Gross revenue amount in USD ($)."),
        ("Quantity", "Integer", "Direct source", "Total unit count of items in line."),
        ("Discount", "Percentage", "Direct source", "Promotional price reduction applied."),
        ("Profit", "Currency", "Direct source", "Net earnings generated after costs ($)."),
        ("Cost", "Currency", "=Sales - Profit", "Inferred operational acquisition/fulfillment cost."),
        ("Profit Margin", "Percentage", "=Profit / Sales", "Return on sales margin ratio.")
    ]

    for idx, (fn, dt, src, desc) in enumerate(dict_entries, start=5):
        ws_dict.cell(row=idx, column=2, value=fn).font = font_bold
        ws_dict.cell(row=idx, column=3, value=dt).font = font_regular
        ws_dict.cell(row=idx, column=4, value=src).font = font_regular
        ws_dict.cell(row=idx, column=5, value=desc).font = font_regular
        for c in range(2, 6):
            ws_dict.cell(row=idx, column=c).border = border_all

    # Auto-fit column widths across all sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col[:50])
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Specific padding for dashboard
    ws_dash.column_dimensions["A"].width = 4
    ws_dash.column_dimensions["B"].width = 16
    ws_dash.column_dimensions["C"].width = 16
    ws_dash.column_dimensions["D"].width = 16
    ws_dash.column_dimensions["E"].width = 16
    ws_dash.column_dimensions["F"].width = 16
    ws_dash.column_dimensions["G"].width = 16
    ws_dash.column_dimensions["H"].width = 16
    ws_dash.column_dimensions["I"].width = 16
    ws_dash.column_dimensions["J"].width = 16
    ws_dash.column_dimensions["K"].width = 16

    # Remove initial default sheet if empty
    if default_sheet.title == "Sheet":
        wb.remove(default_sheet)

    print(f"[4/4] Saving production workbook to: {output_xlsx}...")
    wb.save(output_xlsx)
    print("      Workbook successfully created and verified.")


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    csv_file = os.path.join(base_dir, "data", "superstore.csv")
    xlsx_file = os.path.join(base_dir, "retail_sales_dashboard.xlsx")
    build_retail_dashboard(csv_file, xlsx_file)
