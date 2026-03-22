import pandas as pd

# 1️⃣ Load data
df = pd.read_csv("sales.csv")

# 2️⃣ Fill missing price with 0 (optional) and calculate revenue
df["price"] = df["price"].fillna(0)
df["revenue"] = df["price"] * df["quantity"]

# 3️⃣ Category-wise revenue
category_summary = df.groupby("category")["revenue"].sum().reset_index()

# 4️⃣ Order-wise revenue
order_summary = df.groupby("order_id")["revenue"].sum().reset_index()

# 5️⃣ Create Excel dashboard
with pd.ExcelWriter("final_dashboard.xlsx", engine="xlsxwriter") as writer:
    # Write data
    category_summary.to_excel(writer, sheet_name="Dashboard", startrow=1, index=False)
    order_summary.to_excel(writer, sheet_name="Dashboard", startrow=10, index=False)

    workbook  = writer.book
    worksheet = writer.sheets["Dashboard"]

    # Bar chart: Category-wise Revenue
    bar_chart = workbook.add_chart({"type": "column"})
    bar_chart.add_series({
        "name": "Revenue",
        "categories": f"=Dashboard!A2:A{1 + len(category_summary)}",
        "values": f"=Dashboard!B2:B{1 + len(category_summary)}",
    })
    bar_chart.set_title({"name": "Category-wise Revenue"})
    worksheet.insert_chart("D2", bar_chart)

    # Line chart: Order-wise Revenue
    line_chart = workbook.add_chart({"type": "line"})
    line_chart.add_series({
        "name": "Revenue",
        "categories": f"=Dashboard!A11:A{10 + len(order_summary)}",
        "values": f"=Dashboard!B11:B{10 + len(order_summary)}",
    })
    line_chart.set_title({"name": "Order-wise Revenue Trend"})
    worksheet.insert_chart("D12", line_chart)

print("✅ Dashboard created successfully!")
