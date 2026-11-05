
import csv
import random
import sqlite3
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrData, StrVal
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# SALES & OPERATIONS ANALYTICS PROJECT
# Python + Excel + SQL
# ==========================================

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
OUTPUTS = ROOT / "outputs"

DATA.mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)

rng = random.Random(42)

# PRODUCT CATALOG
products = [
    ("Laptop", "Electronics", 780, 590),
    ("Monitor", "Electronics", 230, 145),
    ("Keyboard", "Accessories", 65, 30),
    ("Mouse", "Accessories", 35, 12),
    ("Desk", "Furniture", 320, 190),
    ("Office Chair", "Furniture", 210, 120),
    ("Headphones", "Electronics", 110, 60),
    ("Webcam", "Accessories", 85, 42),
]

regions = ["North", "South", "East", "West"]

headers = [
    "order_id", "order_date", "customer_id",
    "product", "category", "region",
    "quantity", "unit_price", "unit_cost",
    "discount", "revenue", "profit"
]

# ==========================================
# STEP 1: GENERATE SALES DATA
# ==========================================

rows = []

for i in range(1, 501):
    name, category, price, cost = rng.choice(products)
    qty = rng.randint(1, 6)
    discount = rng.choice([0, 0, 0, 0.05, 0.10, 0.15])

    day = date(2025, 1, 1) + timedelta(
        days=rng.randrange(365)
    )

    revenue = round(qty * price * (1 - discount), 2)
    profit = round(revenue - qty * cost, 2)

    rows.append([
        i,
        day.isoformat(),
        f"C{rng.randint(1, 110):03d}",
        name,
        category,
        rng.choice(regions),
        qty,
        price,
        cost,
        discount,
        revenue,
        profit
    ])

# ==========================================
# STEP 2: EXPORT CSV
# ==========================================

csv_path = DATA / "sales_data.csv"

with csv_path.open("w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    writer.writerows(rows)

# ==========================================
# STEP 3: CREATE SQL DATABASE
# ==========================================

db_path = DATA / "sales.db"
conn = sqlite3.connect(db_path)

conn.execute("DROP TABLE IF EXISTS sales")

conn.execute("""
CREATE TABLE sales (
    order_id INTEGER PRIMARY KEY,
    order_date TEXT,
    customer_id TEXT,
    product TEXT,
    category TEXT,
    region TEXT,
    quantity INTEGER,
    unit_price REAL,
    unit_cost REAL,
    discount REAL,
    revenue REAL,
    profit REAL
)
""")

placeholders = ",".join(["?"] * len(headers))

conn.executemany(
    f"INSERT INTO sales VALUES ({placeholders})",
    rows
)

conn.commit()

# ==========================================
# STEP 4: SQL ANALYSIS
# ==========================================

monthly = conn.execute("""
SELECT
    substr(order_date, 1, 7) AS month,
    COUNT(*) AS orders,
    ROUND(SUM(revenue), 2) AS revenue,
    ROUND(SUM(profit), 2) AS profit
FROM sales
GROUP BY month
ORDER BY month
""").fetchall()

categories = conn.execute("""
SELECT
    category,
    ROUND(SUM(revenue), 2) AS revenue,
    ROUND(SUM(profit), 2) AS profit
FROM sales
GROUP BY category
ORDER BY revenue DESC
""").fetchall()

regional = conn.execute("""
SELECT
    region,
    ROUND(SUM(revenue), 2) AS revenue
FROM sales
GROUP BY region
ORDER BY revenue DESC
""").fetchall()

metrics = conn.execute("""
SELECT
    COUNT(*),
    ROUND(SUM(revenue), 2),
    ROUND(SUM(profit), 2),
    COUNT(DISTINCT customer_id)
FROM sales
""").fetchone()

conn.close()

# ==========================================
# STEP 5: CREATE EXCEL REPORT
# ==========================================

wb = Workbook()

summary = wb.active
summary.title = "Executive Summary"

summary.append(["SALES & OPERATIONS | 2025", "VALUE"])
summary.append(["Total Orders", metrics[0]])
summary.append(["Total Revenue", metrics[1]])
summary.append(["Total Profit", metrics[2]])
summary.append(["Unique Customers", metrics[3]])
summary.append([
    "Profit Margin",
    round(metrics[2] / metrics[1], 4)
])

def add_sheet(name, column_names, records):
    ws = wb.create_sheet(name)
    ws.append(column_names)

    for record in records:
        ws.append(list(record))

    return ws

monthly_ws = add_sheet(
    "Monthly Sales",
    ["Month", "Orders", "Revenue", "Profit"],
    monthly
)

category_ws = add_sheet(
    "Category Performance",
    ["Category", "Revenue", "Profit"],
    categories
)

region_ws = add_sheet(
    "Regional Performance",
    ["Region", "Revenue"],
    regional
)

raw_ws = add_sheet(
    "Raw Sales Data",
    headers,
    rows
)

# ==========================================
# STEP 6: PROFESSIONAL FORMATTING
# ==========================================

for ws in wb.worksheets:
    ws.freeze_panes = "A2"

    if ws != summary:
        ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.fill = PatternFill(
            "solid",
            fgColor="17365D"
        )

        cell.font = Font(
            color="FFFFFF",
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    for col in ws.columns:
        letter = get_column_letter(col[0].column)

        width = max(
            len(str(cell.value or ""))
            for cell in col
        )

        ws.column_dimensions[letter].width = min(
            max(width + 3, 15),
            35
        )

    for row in ws.iter_rows(min_row=2):
        if row[0].row % 2 == 0:
            for cell in row:
                cell.fill = PatternFill(
                    "solid",
                    fgColor="EAF1F8"
                )

# Currency formatting
for cell in ["B3", "B4"]:
    summary[cell].number_format = '$#,##0.00'

summary["B6"].number_format = "0.0%"

for ws, columns in [
    (monthly_ws, [3, 4]),
    (category_ws, [2, 3]),
    (region_ws, [2]),
    (raw_ws, [8, 9, 11, 12])
]:
    for row in ws.iter_rows(min_row=2):
        for index in columns:
            row[index - 1].number_format = '$#,##0.00'

for row in raw_ws.iter_rows(min_row=2):
    row[9].number_format = "0%"

# ==========================================
# STEP 7: MONTHLY REVENUE CHART
# ==========================================

chart = LineChart()

chart.title = "Monthly Revenue Trend"
chart.y_axis.title = "Revenue ($)"
chart.x_axis.title = "Month"

chart.style = 13
chart.width = 22
chart.height = 11
chart.legend = None

# Monthly revenue values
values = Reference(
    monthly_ws,
    min_col=3,
    min_row=1,
    max_row=monthly_ws.max_row
)

chart.add_data(
    values,
    titles_from_data=True,
    from_rows=False
)

# Explicit month labels
months = [
    str(monthly_ws.cell(row=i, column=1).value)
    for i in range(2, monthly_ws.max_row + 1)
]

chart.series[0].cat = AxDataSource(
    strLit=StrData(
        pt=[
            StrVal(idx=i, v=month)
            for i, month in enumerate(months)
        ],
        ptCount=len(months)
    )
)

chart.x_axis.tickLblPos = "nextTo"
chart.y_axis.numFmt = '$#,##0'

chart.series[0].graphicalProperties.line.solidFill = "4472C4"
chart.series[0].graphicalProperties.line.width = 28575
chart.series[0].smooth = False

monthly_ws.add_chart(chart, "F2")

# ==========================================
# STEP 8: SAVE EXCEL REPORT
# ==========================================

excel_path = OUTPUTS / "sales_report.xlsx"

wb.save(excel_path)

# ==========================================
# PROJECT RESULTS
# ==========================================

print("=" * 45)
print("PROJECT COMPLETED SUCCESSFULLY!")
print("=" * 45)

print("500 sales orders generated.")
print("CSV:", csv_path)
print("SQL database:", db_path)
print("Excel report:", excel_path)

print()
print("SALES PERFORMANCE SUMMARY")
print("-" * 35)

print("Total Orders:", metrics[0])
print("Unique Customers:", metrics[3])

print(
    "Total Revenue: ${:,.2f}".format(metrics[1])
)

print(
    "Total Profit: ${:,.2f}".format(metrics[2])
)

print(
    "Profit Margin: {:.1f}%".format(
        metrics[2] / metrics[1] * 100
    )
)

print()
print("READY FOR POWER BI AND GITHUB!")
