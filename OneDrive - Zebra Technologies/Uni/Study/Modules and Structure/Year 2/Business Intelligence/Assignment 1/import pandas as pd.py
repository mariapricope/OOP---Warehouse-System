

#new branch test
#new branch test 2 test-branch

import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import sqlite3

file_path = r"C:\Users\mp3296\OneDrive - Zebra Technologies\Uni\Study\Modules and Structure\Year 2\Business Intelligence\Assignment 1\Data Sets for Assignment One.xlsx"  
excel_data = pd.ExcelFile(file_path)

sales_data = excel_data.parse("Sales Data")
customer_data = excel_data.parse("Customer Data")
marketing_data = excel_data.parse("Marketing Campaign Data")

print("\nSales Data :\n", sales_data.head())
print("\nCustomer Data :\n", customer_data.head())
print("\nMarketing Data :\n", marketing_data.head())

sales_data['Product Name'] = sales_data['Product Name'].str.lower()

print("\nChecking for missing values:")
print("Missing values in 'Region':", customer_data['Region'].isna().sum())
print("Missing values in 'Quantity Sold':", sales_data['Quantity Sold'].isna().sum())
print("Missing values in 'Start Date':", marketing_data['Start Date'].isna().sum())
print("Missing values in 'End Date':", marketing_data['End Date'].isna().sum())



customer_data['Region'] = customer_data['Region'].fillna('Unknown')
sales_data.dropna(subset=['Product Name', 'Quantity Sold'], inplace=True)
marketing_data.dropna(subset =['Start Date', 'End Date'], inplace =True)

print("\nRemoving duplicates in Sales, Customer, and Marketing Data:")

sales_data.drop_duplicates(inplace=True)
customer_data.drop_duplicates(inplace=True)
marketing_data.drop_duplicates(inplace=True)

print("\nConverting 'Quantity Sold' to int and 'Revenue (USD)' to float:")

sales_data['Quantity Sold'] = sales_data['Quantity Sold'].astype(int)
sales_data['Revenue (USD)'] = sales_data['Revenue (USD)'].replace(r'[\$,]', '', regex=True).astype(float)
customer_data['Total Revenue (USD)'] = customer_data['Total Revenue (USD)'].replace(r'[\$,]', '', regex=True).astype(float)

print("\nNormalising 'Quantity Sold' data:")

sales_data['Quantity_Scaled'] = (sales_data['Quantity Sold'] - sales_data['Quantity Sold'].min()) / (sales_data['Quantity Sold'].max() - sales_data['Quantity Sold'].min())

print(sales_data['Quantity_Scaled'])

print("\nConverting 'Date' to datetime format:\n", sales_data)
sales_data['Date'] = pd.to_datetime(sales_data['Date'], errors='coerce')
marketing_data['Start Date'] = pd.to_datetime(marketing_data['Start Date'], errors='coerce')
marketing_data['End Date'] = pd.to_datetime(marketing_data['End Date'], errors='coerce')



print("\nCheck for NaNs in Quantity Sold:\n", sales_data['Quantity Sold'].isna().sum())
print("\nCheck for rows where Quantity Sold is zero:\n", (sales_data['Quantity Sold'] == 0).sum())
print("\nSummary statistics of Quantity Sold:\n", sales_data['Quantity Sold'].describe())

print("\nAggregating sales data by month:")
sales_data['Month'] = sales_data['Date'].dt.to_period('M').astype(str)
monthly_sales = sales_data.groupby('Month').agg(
    Total_Sold=('Quantity Sold', 'sum'),
    Total_Revenue=('Revenue (USD)', 'sum'),
    Average_Sold=('Quantity Sold', 'mean')
).reset_index()

print("\nMonthly Sales Aggregation:\n", monthly_sales)


# Step 1: Split the 'Products Promoted' column by commas in the marketing_data DataFrame
# Step 2: Explode the 'Products Promoted' column so each product is in a separate row
# Step 3: Clean up the extra spaces (if any) in the 'Products Promoted' column  

marketing_data['Products Promoted'] = marketing_data['Products Promoted'].str.split(',')
marketing_data = marketing_data.explode('Products Promoted')
marketing_data['Products Promoted'] = marketing_data['Products Promoted'].str.strip()



print("\nCreating temporary SQL database:")

conn = sqlite3.connect(':memory:')

sales_data.to_sql('sales_data', conn, if_exists='replace', index=False)
customer_data.to_sql('customer_data', conn, if_exists='replace', index=False)
marketing_data.to_sql('marketing_data', conn, if_exists='replace', index=False)
monthly_sales.to_sql('monthly_sales', conn, if_exists='replace', index=False)

print("\nRunning SQL query to get the top 5 most sold products:")
query_top_products = """
SELECT "Product Name", SUM("Quantity Sold") AS Total_Sold
FROM sales_data
GROUP BY "Product Name"
ORDER BY Total_Sold DESC
LIMIT 5;
"""
top_products = pd.read_sql(query_top_products, conn)
print("Top 5 Most Sold Products:\n", top_products)

print("\nRunning SQL query to get total purchases by region:")
query_sales_by_region_purchases = """
SELECT Region, SUM("Total Purchases") AS Total_Purchases
FROM customer_data
GROUP BY Region
ORDER BY Total_Purchases DESC;
"""
sales_by_region_purchases = pd.read_sql(query_sales_by_region_purchases, conn)
print("\nTotal Purchases by Region:\n", sales_by_region_purchases)

if not sales_by_region_purchases.empty and sales_by_region_purchases['Total_Purchases'].sum() > 0:
    print("\nPlotting Purchases Distribution by Region:")
    plt.figure(figsize=(8, 5))
    plt.pie(
        sales_by_region_purchases['Total_Purchases'],
        labels=sales_by_region_purchases['Region'],
        autopct='%1.1f%%',
        startangle=140
    )
    plt.title('Purchases Distribution by Region')
    plt.tight_layout()
    plt.show()
else:
    print("No purchase data available for regions, or Total_Purchases is zero.")

print("\nRunning SQL query to get sales impact by marketing campaign:")


# Step 5: Run the SQL query to get the sales impact by marketing campaign
query_marketing_sales = """
SELECT 
    marketing_data."Campaign Name", 
    sales_data."Product Category",
    SUM(sales_data."Quantity Sold") AS Sales_During_Campaign
FROM 
    sales_data
JOIN 
    marketing_data
    ON sales_data."Product Category" = marketing_data."Products Promoted" 
    AND sales_data."Date" BETWEEN marketing_data."Start Date" AND marketing_data."End Date"
GROUP BY 
    marketing_data."Campaign Name", sales_data."Product Category"
ORDER BY 
    Sales_During_Campaign DESC;
"""

marketing_sales = pd.read_sql(query_marketing_sales, conn)
print("\nSales Impact by Marketing Campaign:\n", marketing_sales)

print("\nRunning SQL query to check which discount attracted more sales:")
query_sales_by_discount = """
SELECT 
    marketing_data."Discount Offered (%)", 
    SUM(sales_data."Quantity Sold") AS Total_Sold,
    SUM(sales_data."Revenue (USD)") AS Total_Revenue
FROM 
    sales_data
JOIN 
    marketing_data 
    ON sales_data."Product Category" = marketing_data."Products Promoted" 
    AND sales_data."Date" BETWEEN marketing_data."Start Date" AND marketing_data."End Date"
GROUP BY 
    marketing_data."Discount Offered (%)"
ORDER BY 
    Total_Sold DESC;
"""

# Execute the query and fetch results
sales_by_discount = pd.read_sql(query_sales_by_discount, conn)

# Print the result
print("\nSales Impact by Discount Offered:\n", sales_by_discount)

if not sales_by_discount.empty:
    plt.figure(figsize=(8, 5))
    sns.barplot(x='Discount Offered (%)', y='Total_Sold', data=sales_by_discount, palette='viridis')
    plt.title('Sales by Discount Percentage')
    plt.xlabel('Discount Offered (%)')
    plt.ylabel('Total Sold')
    plt.tight_layout()
    plt.show()
else:
    print("No sales data available for the specified discount periods.")

print("\nRunning SQL query to check which customer spent the most and from which region:")
query_spent_most_region = """
SELECT "Customer Name", 
       "Region", 
       SUM("Total Revenue (USD)") AS Total_Spent
FROM customer_data
GROUP BY "Customer Name", "Region"
ORDER BY Total_Spent DESC
LIMIT 1;
"""
spent_most_region = pd.read_sql(query_spent_most_region, conn)
print("\nCustomer who spent the most and their region:\n", spent_most_region)

#Retention Rates

end_date = sales_data['Date'].max()
start_date = end_date - pd.DateOffset(years=1)

last_year_sales = sales_data[(sales_data['Date'] >= start_date) & (sales_data['Date'] <= end_date)]

mid_date = start_date + pd.DateOffset(months=6)

first_period_customers = last_year_sales[last_year_sales['Date'] < mid_date]['Order ID'].unique()
second_period_customers = last_year_sales[last_year_sales['Date'] >= mid_date]['Order ID'].unique()

initial_customers = set(first_period_customers)
returning_customers = initial_customers.intersection(second_period_customers)

retention_rate = len(returning_customers) / len(initial_customers) if len(initial_customers) > 0 else 0

print(f"Customer Retention Rate: {retention_rate:.2%}")


sns.set(style="whitegrid")

print("\nCheck Data Types of 'Total_Sold' and 'Total_Revenue':")
print(monthly_sales[['Total_Sold', 'Total_Revenue']].dtypes)

print("\nCheck for NaN values in Total Sold and Total Revenue:")
print(monthly_sales[['Total_Sold', 'Total_Revenue']].isna().sum())

print("\nCheck for zero values in 'Total_Sold' and 'Total_Revenue':")
print((monthly_sales[['Total_Sold', 'Total_Revenue']] == 0).sum())

fig, ax1 = plt.subplots(figsize=(10, 6))

sns.lineplot(x=monthly_sales['Month'].astype(str), y=monthly_sales['Total_Sold'], label='Total Sold', color='b', marker='o', ax=ax1)
ax1.set_xlabel('Month')
ax1.set_ylabel('Total Sold', color='b', labelpad=15)
ax1.tick_params(axis='y', labelcolor='b')

ax2 = ax1.twinx()
sns.lineplot(x=monthly_sales['Month'].astype(str), y=monthly_sales['Total_Revenue'], label='Total Revenue', color='r', marker='o', ax=ax2)
ax2.set_ylabel('Total Revenue (USD)', color='r', labelpad=15)
ax2.tick_params(axis='y', labelcolor='r')

plt.title('Monthly Sales Trend')
plt.xticks(rotation=45)
ax1.legend(loc='upper left', bbox_to_anchor=(0, 1.1))
ax2.legend(loc='upper right', bbox_to_anchor=(1, 1.1))
plt.tight_layout()
plt.show()

print("\nCheck first few rows of monthly_sales to confirm data before plotting:")
print(monthly_sales.head())

print("\nPlotting Top 5 Most Sold Products:")
plt.figure(figsize=(8, 5))
sns.barplot(x='Product Name', y='Total_Sold', data=top_products, palette="viridis", hue='Product Name', legend=False)
plt.title('Top 5 Most Sold Products')
plt.xlabel('Product Name')
plt.ylabel('Total Sold')
plt.tight_layout()
plt.show()

print("\nClosing SQL connection.")
conn.close()
