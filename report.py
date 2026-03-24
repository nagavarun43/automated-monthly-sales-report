from pathlib import Path
from datetime import datetime
import pandas as pd

# Get project folder
base_path = Path(__file__).resolve().parent

# Input file
file_path = base_path / "sales.xlsx"

print("Looking for file:", file_path)

if not file_path.exists():
    raise FileNotFoundError(f"sales.xlsx not found in {base_path}")

# Read Excel file
df = pd.read_excel(file_path, engine="openpyxl")

# Clean column names
df.columns = df.columns.str.strip()

# Convert ORDERDATE to datetime
df["ORDERDATE"] = pd.to_datetime(df["ORDERDATE"], errors="coerce")

# Convert numeric columns
numeric_cols = [
    "SALES",
    "QUANTITYORDERED",
    "PRICEEACH",
    "MSRP",
    "MONTH_ID",
    "YEAR_ID",
    "QTR_ID"
]

for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Drop rows where important fields are missing
df = df.dropna(subset=["ORDERDATE", "SALES"])

# Create Month column like 2003-02
df["Month"] = df["ORDERDATE"].dt.to_period("M").astype(str)

# Monthly summary
monthly_summary = df.groupby("Month").agg(
    Total_Sales=("SALES", "sum"),
    Total_Orders=("ORDERNUMBER", "nunique"),
    Total_Quantity=("QUANTITYORDERED", "sum"),
    Avg_Order_Value=("SALES", "mean")
).reset_index()

monthly_summary["Total_Sales"] = monthly_summary["Total_Sales"].round(2)
monthly_summary["Avg_Order_Value"] = monthly_summary["Avg_Order_Value"].round(2)

# Product line summary
productline_summary = df.groupby(["Month", "PRODUCTLINE"]).agg(
    Total_Sales=("SALES", "sum"),
    Total_Quantity=("QUANTITYORDERED", "sum")
).reset_index()

productline_summary["Total_Sales"] = productline_summary["Total_Sales"].round(2)

# Country summary
country_summary = df.groupby(["Month", "COUNTRY"]).agg(
    Total_Sales=("SALES", "sum"),
    Total_Orders=("ORDERNUMBER", "nunique")
).reset_index()

country_summary["Total_Sales"] = country_summary["Total_Sales"].round(2)

# Insights
top_month = monthly_summary.loc[monthly_summary["Total_Sales"].idxmax()]
top_month_name = top_month["Month"]
top_month_sales = top_month["Total_Sales"]

top_product = df.groupby("PRODUCTLINE")["SALES"].sum().idxmax()
top_product_sales = round(df.groupby("PRODUCTLINE")["SALES"].sum().max(), 2)

top_country = df.groupby("COUNTRY")["SALES"].sum().idxmax()
top_country_sales = round(df.groupby("COUNTRY")["SALES"].sum().max(), 2)

insights = pd.DataFrame({
    "Insight": [
        f"Highest sales month: {top_month_name} with total sales of {top_month_sales}",
        f"Top product line overall: {top_product} with total sales of {top_product_sales}",
        f"Top country overall: {top_country} with total sales of {top_country_sales}"
    ]
})

# Output file with timestamp to avoid permission issues
output_path = base_path / f"monthly_sales_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# Save Excel report
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    monthly_summary.to_excel(writer, sheet_name="Monthly_Summary", index=False)
    productline_summary.to_excel(writer, sheet_name="ProductLine_Summary", index=False)
    country_summary.to_excel(writer, sheet_name="Country_Summary", index=False)
    insights.to_excel(writer, sheet_name="Insights", index=False)

print(f"Report created successfully: {output_path.name}")
print("\nMonthly Summary Preview:")
print(monthly_summary.head())