import pandas as pd

df = pd.read_excel("sales_data.xlsx")

df["Date"] = pd.to_datetime(df["Date"])

df["TotalSales"] = df["Quantity"] * df["Price"]

sales_by_product = (
    df.groupby("Product")["TotalSales"]
    .sum()
    .sort_values(ascending=False)
)

sales_by_region = (
    df.groupby("Region")["TotalSales"]
    .sum()
    .sort_values()
)

df["Month"] = df["Date"].dt.to_period("M")
monthly_sales = df.groupby("Month")["TotalSales"].sum()

best_product = sales_by_product.idxmax()
worst_region = sales_by_region.idxmin()

summary = pd.DataFrame({
    "Metric": [
        "Best Selling Product",
        "Worst Performing Region",
        "Total Revenue"
    ],
    "Value": [
        best_product,
        worst_region,
        df["TotalSales"].sum()
    ]
})

summary.to_excel("sales_summary.xlsx", index=False)

print("Sales performance summary")
print("\nSales by Product:\n", sales_by_product)
print("\nSales by Region:\n", sales_by_region)
print("\nMonthly Sales:\n", monthly_sales)
print("\nSummary:\n", summary)
print("Sales performance summary saved to sales_summary.xlsx")