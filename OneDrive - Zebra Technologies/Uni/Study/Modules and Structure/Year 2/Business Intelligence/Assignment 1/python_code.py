# Importing Libraries
import pandas as pd  # For data manipulation and analysis
import numpy as np   # For numerical operations
import seaborn as sns  # For statistical data visualisation
import matplotlib.pyplot as plt  # For plotting graphs
import sqlite3  # For database operations

# --------------------------- Data Loading Section ---------------------------
# Load the Excel file from path
file_path = r"C:\Users\mp3296\OneDrive - Zebra Technologies\Uni\Study\Modules and Structure\Year 2\Business Intelligence\Assignment 1\Data Sets for Assignment One.xlsx"  
excel_data = pd.ExcelFile(file_path)

# Load individual sheets into DataFrames
sales_data = excel_data.parse("Sales Data")
customer_data = excel_data.parse("Customer Data")
marketing_data = excel_data.parse("Marketing Campaign Data")

# Debugging: Checking the first few rows of each DataFrame to ensure they loaded correctly
print("\nSales Data :\n", sales_data.head())
print("\nCustomer Data :\n", customer_data.head())
print("\nMarketing Data :\n", marketing_data.head())

# --------------------------- Data Cleaning Section ---------------------------

# Step 1: Standardising Product Names
print("\nStandardising Product Names:")
sales_data['Product Name'] = sales_data['Product Name'].str.lower()  # Convert all product names to lowercase for consistency

# Step 2: Handle Missing Values
print("\nChecking for missing values in 'Region' and 'Quantity Sold' columns:")
print("Missing values in 'Region':", customer_data['Region'].isna().sum())  # Check for missing regions
print("Missing values in 'Quantity Sold':", sales_data['Quantity Sold'].isna().sum())  # Check for missing quantities sold

# Fill missing region values with 'Unknown' and remove rows with missing product name or quantity sold
customer_data['Region'] = customer_data['Region'].fillna('Unknown')  # Impute missing region with 'Unknown'
sales_data.dropna(subset=['Product Name', 'Quantity Sold'], inplace=True)  # Drop rows with missing values in essential columns

# Step 3: Removing Duplicates
print("\nRemoving duplicates in Sales, Customer, and Marketing Data:")
sales_data.drop_duplicates(inplace=True)  # Remove duplicate rows from sales data
customer_data.drop_duplicates(inplace=True)  # Remove duplicate rows from customer data
marketing_data.drop_duplicates(inplace=True)  # Remove duplicate rows from marketing data

# Step 4: Converting Data Types
print("\nConverting 'Quantity Sold' to int and 'Revenue (USD)' to float:")
sales_data['Quantity Sold'] = sales_data['Quantity Sold'].astype(int)  # Convert quantity sold to integer
sales_data['Revenue (USD)'] = sales_data['Revenue (USD)'].replace(r'[\$,]', '', regex=True).astype(float)  # Remove dollar signs and convert to float

# Step 5: Data Normalisation
print("\nNormalising 'Quantity Sold' data:")
sales_data['Quantity_Scaled'] = (sales_data['Quantity Sold'] - sales_data['Quantity Sold'].min()) / (sales_data['Quantity Sold'].max() - sales_data['Quantity Sold'].min())  # Scale quantity sold

# Step 6: Standardising Dates
print("\nConverting 'Date' to datetime format:")
sales_data['Date'] = pd.to_datetime(sales_data['Date'], errors='coerce')
marketing_data['Start Date'] = pd.to_datetime(marketing_data['Start Date'], errors='coerce')
marketing_data['End Date'] = pd.to_datetime(marketing_data['End Date'], errors='coerce')
 # Convert 'Date' column to datetime format

# Step 7: Checking for NaNs and Zero Values
print("\nCheck for NaNs in Quantity Sold:\n", sales_data['Quantity Sold'].isna().sum())  # Check for NaN values in Quantity Sold
print("\nCheck for rows where Quantity Sold is zero:\n", (sales_data['Quantity Sold'] == 0).sum())  # Check for zero values in Quantity Sold
print("\nSummary statistics of Quantity Sold:\n", sales_data['Quantity Sold'].describe())  # Summary statistics for Quantity Sold

# --------------------------- Data Aggregation Section ---------------------------
# Step 8: Aggregate Daily Sales into Monthly Sales
print("\nAggregating sales data by month:")
sales_data['Month'] = sales_data['Date'].dt.to_period('M').astype(str)  # Extract month from date
monthly_sales = sales_data.groupby('Month').agg(
    Total_Sold=('Quantity Sold', 'sum'),  # Total quantity sold by month
    Total_Revenue=('Revenue (USD)', 'sum'),  # Total revenue by month
    Average_Sold=('Quantity Sold', 'mean')  # Average quantity sold by month
).reset_index()

print("\nMonthly Sales Aggregation:\n", monthly_sales)  # Display the aggregated monthly sales data

# --------------------------- Temporary SQL Database Creation ---------------------------
# Creating an in-memory SQL database
print("\nCreating temporary SQL database:")
conn = sqlite3.connect(':memory:')  # In-memory temporary database

# Storing DataFrames into SQL tables for querying
sales_data.to_sql('sales_data', conn, if_exists='replace', index=False)  # Sales data table
customer_data.to_sql('customer_data', conn, if_exists='replace', index=False)  # Customer data table
marketing_data.to_sql('marketing_data', conn, if_exists='replace', index=False)  # Marketing data table
monthly_sales.to_sql('monthly_sales', conn, if_exists='replace', index=False)  # Aggregated monthly sales table

# --------------------------- SQL Queries for Insights ---------------------------

# Query 1: Top 5 Most Sold Products
print("\nRunning SQL query to get the top 5 most sold products:")
query_top_products = """
SELECT "Product Name", 
SUM("Quantity Sold") AS Total_Sold, 
SUM("Revenue (USD)") AS Total_Revenue
FROM sales_data
GROUP BY "Product Name"
ORDER BY Total_Revenue DESC
LIMIT 5;
"""
top_products = pd.read_sql(query_top_products, conn)
print("Top 5 Most Sold Products:\n", top_products)  # Display the top 5 most sold products

# Query 2: Total Purchases by Region
print("\nRunning SQL query to get total purchases by region:")
query_sales_by_region_purchases = """
SELECT Region, SUM("Total Purchases") AS Total_Purchases
FROM customer_data
GROUP BY Region
ORDER BY Total_Purchases DESC;
"""
sales_by_region_purchases = pd.read_sql(query_sales_by_region_purchases, conn)
print("\nTotal Purchases by Region:\n", sales_by_region_purchases)  # Display total purchases by region

# Plot purchases distribution by region (if data is available)
if not sales_by_region_purchases.empty and sales_by_region_purchases['Total_Purchases'].sum() > 0:
    print("\nPlotting Purchases Distribution by Region:")
    plt.figure(figsize=(8, 5))
    plt.pie(
        sales_by_region_purchases['Total_Purchases'],
        labels=sales_by_region_purchases['Region'],
        autopct='%1.1f%%',
        startangle=140
    )
    plt.title('Purchases Distribution by Region')  # Pie chart title
    plt.tight_layout()
    plt.show()
else:
    print("No purchase data available for regions, or Total_Purchases is zero.")

# Query 3: Marketing Campaign Impact
print("\nRunning SQL query to get sales impact by marketing campaign:")
query_marketing_sales = """
SELECT marketing_data."Campaign Name", SUM(sales_data."Quantity Sold") AS Sales_During_Campaign
FROM sales_data
JOIN marketing_data ON sales_data."Date" BETWEEN marketing_data."Start Date" AND marketing_data."End Date"
GROUP BY marketing_data."Campaign Name"
ORDER BY Sales_During_Campaign DESC;
"""
marketing_sales = pd.read_sql(query_marketing_sales, conn)
print("\nSales Impact by Marketing Campaign:\n", marketing_sales)  # Display marketing campaign impact

# Query 4: Sales Impact by Discount Offered
print("\nRunning SQL query to check which discount attracted more sales:")
query_sales_by_discount = """
SELECT marketing_data."Discount Offered (%)", 
SUM(sales_data."Quantity Sold") AS Total_Sold,
SUM(sales_data."Revenue (USD)") AS Total_Revenue
FROM sales_data
JOIN marketing_data ON sales_data."Date" BETWEEN marketing_data."Start Date" AND marketing_data."End Date"
GROUP BY marketing_data."Discount Offered (%)"
ORDER BY Total_Sold DESC;
"""
sales_by_discount = pd.read_sql(query_sales_by_discount, conn)
print("\nSales Impact by Discount Offered:\n", sales_by_discount)  # Display sales by discount percentage

# If the result is valid, plot the data
if not sales_by_discount.empty:
    plt.figure(figsize=(8, 5))
    sns.barplot(x='Discount Offered (%)', y='Total_Sold', data=sales_by_discount, palette='viridis')
    plt.title('Sales by Discount Percentage')  # Bar plot title
    plt.xlabel('Discount Offered (%)')
    plt.ylabel('Total Sold')
    plt.tight_layout()
    plt.show()
else:
    print("No sales data available for the specified discount periods.")

# Query 7: Check which customer spent the most and from which region
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
print("\nCustomer who spent the most and their region:\n", spent_most_region)  # Display the customer who spent the most

# Query 8: Caluclate retention rates:


# Define the time period for the past year
end_date = sales_data['Date'].max()  # Latest date in the sales data
start_date = end_date - pd.DateOffset(years=1)

# Filter sales data to only include transactions from the last year
last_year_sales = sales_data[(sales_data['Date'] >= start_date) & (sales_data['Date'] <= end_date)]

# Split the last year into two 6-month periods
mid_date = start_date + pd.DateOffset(months=6)

# Get unique customer IDs for each period
first_period_customers = last_year_sales[last_year_sales['Date'] < mid_date]['Order ID'].unique()
second_period_customers = last_year_sales[last_year_sales['Date'] >= mid_date]['Order ID'].unique()

# Convert to sets for easy comparison
initial_customers = set(first_period_customers)
returning_customers = initial_customers.intersection(second_period_customers)

# Calculate retention rate
retention_rate = len(returning_customers) / len(initial_customers) if len(initial_customers) > 0 else 0

print(f"Customer Retention Rate: {retention_rate:.2%}")



# --------------------------- Data Visualisation Section ---------------------------

# Set the visual style for Seaborn
sns.set(style="whitegrid")

# Check the data types and any potential NaNs or unexpected values in monthly_sales
print("\nCheck Data Types of 'Total_Sold' and 'Total_Revenue':")
print(monthly_sales[['Total_Sold', 'Total_Revenue']].dtypes)

# Check for NaN values in the monthly_sales DataFrame
print("\nCheck for NaN values in Total Sold and Total Revenue:")
print(monthly_sales[['Total_Sold', 'Total_Revenue']].isna().sum())

# Check for zero values in the monthly sales data
print("\nCheck for zero values in 'Total_Sold' and 'Total_Revenue':")
print((monthly_sales[['Total_Sold', 'Total_Revenue']] == 0).sum())

# Plot Monthly Sales and Revenue Trend
fig, ax1 = plt.subplots(figsize=(10, 6))

# Plot Total Sold on the primary y-axis (left axis)
sns.lineplot(x=monthly_sales['Month'].astype(str), y=monthly_sales['Total_Sold'], label='Total Sold', color='b', marker='o', ax=ax1)
ax1.set_xlabel('Month')
ax1.set_ylabel('Total Sold', color='b', labelpad=15)
ax1.tick_params(axis='y', labelcolor='b')

# Create a secondary y-axis for Total Revenue
ax2 = ax1.twinx()
sns.lineplot(x=monthly_sales['Month'].astype(str), y=monthly_sales['Total_Revenue'], label='Total Revenue', color='r', marker='o', ax=ax2)
ax2.set_ylabel('Total Revenue (USD)', color='r', labelpad=15)
ax2.tick_params(axis='y', labelcolor='r')

# Adding titles and labels
plt.title('Monthly Sales Trend')
plt.xticks(rotation=45)

# Show legends for both lines
ax1.legend(loc='upper left', bbox_to_anchor=(0, 1.1))
ax2.legend(loc='upper right', bbox_to_anchor=(1, 1.1))

# Adjust layout to avoid clipping
plt.tight_layout()

# Display the plot
plt.show()

# Optional: Check first few rows of monthly_sales before plotting
print("\nCheck first few rows of monthly_sales to confirm data before plotting:")
print(monthly_sales.head())

# --------------------------- Plotting Top 5 Most Sold Products ---------------------------
print("\nPlotting Top 5 Most Sold Products:")
plt.figure(figsize=(8, 5))
sns.barplot(x='Product Name', y='Total_Sold', data=top_products, palette="viridis", hue='Product Name', legend=False)
plt.title('Top 5 Most Sold Products')
plt.xlabel('Product Name')
plt.ylabel('Total Sold')
plt.tight_layout()
plt.show()


# --------------------------- Closing SQL Connection ---------------------------
print("\nClosing SQL connection.")
conn.close()  # Close the in-memory database connection
