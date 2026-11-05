
import streamlit as st
import pandas as pd
from pathlib import Path

# ==========================================
# SALES ANALYTICS DASHBOARD
# ==========================================

st.set_page_config(
    page_title="Sales Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

# Load sales dataset
ROOT = Path(__file__).resolve().parent
DATA_PATH = ROOT / "data" / "sales_data.csv"

@st.cache_data
def load_data():
    df = pd.read_csv(DATA_PATH)
    df["order_date"] = pd.to_datetime(df["order_date"])
    return df

if not DATA_PATH.exists():
    st.error("Sales dataset not found. Run main.py first.")
    st.stop()

df = load_data()

# ==========================================
# DASHBOARD HEADER
# ==========================================

st.title("📊 Sales & Operations Analytics")
st.caption("2025 Sales Performance | Python + SQL + Streamlit")
st.info("Portfolio demonstration using synthetic sales data.")

# ==========================================
# SIDEBAR FILTERS
# ==========================================

st.sidebar.header("Dashboard Filters")

selected_regions = st.sidebar.multiselect(
    "Select Region",
    sorted(df["region"].unique()),
    default=sorted(df["region"].unique())
)

selected_categories = st.sidebar.multiselect(
    "Select Category",
    sorted(df["category"].unique()),
    default=sorted(df["category"].unique())
)

filtered = df[
    df["region"].isin(selected_regions)
    & df["category"].isin(selected_categories)
].copy()

# ==========================================
# KEY PERFORMANCE INDICATORS
# ==========================================

total_revenue = filtered["revenue"].sum()
total_profit = filtered["profit"].sum()
total_orders = len(filtered)
unique_customers = filtered["customer_id"].nunique()

profit_margin = (
    total_profit / total_revenue * 100
    if total_revenue != 0 else 0
)

st.subheader("Executive Summary")

col1, col2, col3, col4 = st.columns(4)

col1.metric("Total Revenue", f"${total_revenue:,.2f}")
col2.metric("Total Profit", f"${total_profit:,.2f}")
col3.metric("Total Orders", f"{total_orders:,}")
col4.metric("Profit Margin", f"{profit_margin:.1f}%")

st.metric("Unique Customers", unique_customers)

st.divider()

# ==========================================
# MONTHLY SALES TREND
# ==========================================

st.subheader("Monthly Revenue Trend")

monthly = (
    filtered.assign(
        month=filtered["order_date"].dt.to_period("M").dt.to_timestamp()
    )
    .groupby("month", as_index=False)
    .agg(
        revenue=("revenue", "sum"),
        profit=("profit", "sum"),
        orders=("order_id", "count")
    )
)

st.line_chart(
    monthly.set_index("month")[["revenue"]],
    y_label="Revenue ($)",
    x_label="Month"
)

# ==========================================
# CATEGORY & REGIONAL PERFORMANCE
# ==========================================

left, right = st.columns(2)

with left:
    st.subheader("Revenue by Category")

    category_sales = (
        filtered.groupby("category")["revenue"]
        .sum()
        .sort_values(ascending=False)
    )

    st.bar_chart(
        category_sales,
        y_label="Revenue ($)",
        x_label="Category"
    )

with right:
    st.subheader("Revenue by Region")

    region_sales = (
        filtered.groupby("region")["revenue"]
        .sum()
        .sort_values(ascending=False)
    )

    st.bar_chart(
        region_sales,
        y_label="Revenue ($)",
        x_label="Region"
    )

# ==========================================
# SALES DATA TABLE
# ==========================================

st.divider()
st.subheader("Sales Transactions")

st.dataframe(
    filtered,
    use_container_width=True,
    hide_index=True
)

# ==========================================
# DOWNLOAD FILTERED DATA
# ==========================================

csv_data = filtered.to_csv(index=False).encode("utf-8")

st.download_button(
    label="Download Filtered Sales Data",
    data=csv_data,
    file_name="filtered_sales.csv",
    mime="text/csv"
)

st.caption(
    "Sales & Operations Analytics Portfolio Project | "
    "Built with Python, SQL, Excel and Streamlit"
)

# BUSINESS INSIGHTS
st.divider()
st.header("Business Insights & Recommendations")

if not filtered.empty:
    category_sales = (
        filtered.groupby("category")["revenue"]
        .sum()
        .sort_values(ascending=False)
    )

    region_sales = (
        filtered.groupby("region")["revenue"]
        .sum()
        .sort_values(ascending=False)
    )

    monthly_sales = (
        filtered.assign(
            month=pd.to_datetime(filtered["order_date"])
            .dt.to_period("M")
            .astype(str)
        )
        .groupby("month")["revenue"]
        .sum()
        .sort_values(ascending=False)
    )

    best_category = category_sales.index[0]
    best_region = region_sales.index[0]
    best_month = monthly_sales.index[0]

    st.subheader("Key Findings")

    st.success(
        f"Top-performing category: {best_category} "
        f"(${category_sales.iloc[0]:,.2f} in revenue)"
    )

    st.info(
        f"Highest-revenue region: {best_region} "
        f"(${region_sales.iloc[0]:,.2f} in revenue)"
    )

    st.info(
        f"Best sales month: {best_month} "
        f"(${monthly_sales.iloc[0]:,.2f} in revenue)"
    )

    st.subheader("Business Recommendations")

    st.write(
        f"1. Analyze the products within {best_category} "
        "to identify opportunities for profitable growth."
    )

    st.write(
        f"2. Investigate the factors contributing to "
        f"{best_region}'s sales performance."
    )

    st.write(
        f"3. Examine sales patterns during {best_month} "
        "to identify potential seasonal opportunities."
    )

    st.caption(
        "Insights are based on synthetic portfolio data "
        "and update with the dashboard filters."
    )
else:
    st.warning("No sales data available for the selected filters.")
