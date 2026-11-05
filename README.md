# Sales & Operations Analytics Dashboard
## Live Dashboard

[View Interactive Sales Analytics Dashboard](https://sales-operations-analytics-xmhkedzqw6usetxp4tscp2.streamlit.app/)

Explore revenue, profit, regional performance, product categories, and business insights using 500 synthetic sales transactions.

## Project Overview
An end-to-end sales analytics portfolio project built using Python, SQL, Excel, and Streamlit. The project analyzes 500 synthetic sales transactions from 2025 and presents interactive business insights.

## Technologies Used
- Python
- Pandas
- SQLite / SQL
- Excel (openpyxl)
- Streamlit
- Altair

## Key Features
- Automated generation of 500 synthetic sales orders
- SQLite database and CSV export
- Formatted Excel sales report
- Interactive dashboard with region and category filters
- Revenue, profit, orders, profit margin, and customer KPIs
- Monthly revenue and category/region performance charts
- Downloadable filtered sales data
- Dynamic business insights and recommendations

## Project Results
- Total Orders: 500
- Total Revenue: $352,598.50
- Total Profit: $115,237.50
- Profit Margin: 32.7%
- Unique Customers: 108

These results are based on simulated data and are not actual business performance.

## How to Run
1. Clone or download this repository.
2. Install dependencies: `pip install -r requirements.txt`
3. Generate the sales data: `python main.py`
4. Launch the dashboard: `streamlit run dashboard.py`

## Project Structure
- `main.py` — Generates sales data, SQLite database, and Excel report
- `dashboard.py` — Interactive Streamlit dashboard
- `requirements.txt` — Python dependencies
- `data/` — Generated CSV and SQLite database
- `outputs/` — Generated Excel report

## Business Value
Demonstrates an end-to-end analytics workflow, including data generation, SQL aggregation, reporting, interactive visualization, and data-driven business recommendations.
