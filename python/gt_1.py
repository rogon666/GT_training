###############################################################################
# Cleaning and Consolidating Sales Data from Excel Files
# This script consolidates sales data from two Excel files, cleans the data, 
# and saves the consolidated result to a new Excel file.
###############################################################################
import os
import pandas as pd

# Check current directory
print("Current Directory:", os.getcwd())
# Change to a new directory (example path)
new_dir = "C:/2025/06_june/GT/excel"
os.chdir(new_dir)
# Confirm change
print("New Directory:", os.getcwd())

# Load the Excel files
df1 = pd.read_excel("monthly_sales_v3.xlsx")  # or sheet_name=None for all sheets
df2 = pd.read_excel("monthly_sales_v4 - Copy final use this.xlsx")
df1.head()
df2.head()

# ------------------------------------------------------------
# Rename columns in df2 to match df1
df2_cleaned = df2.rename(columns={
    'code': 'Product',
    'units': 'Units Sold',
    'Price ($)': 'Unit Price ($)',
    'Total ($)': 'Total Sales ($)'
})

# Convert 'Total Sales ($)' to proper numeric format (handle comma as decimal/thousand separator)
df2_cleaned['Total Sales ($)'] = df2_cleaned['Total Sales ($)'].replace({',': ''}, regex=True).astype(float)

# Ensure 'Units Sold' and 'Unit Price ($)' are numeric
df2_cleaned['Units Sold'] = pd.to_numeric(df2_cleaned['Units Sold'], errors='coerce')
df2_cleaned['Unit Price ($)'] = pd.to_numeric(df2_cleaned['Unit Price ($)'], errors='coerce')

# Standardize date format
df1['Date'] = pd.to_datetime(df1['Date'])
df2_cleaned['Date'] = pd.to_datetime(df2_cleaned['Date'])

# Combine the dataframes
consolidated_df = pd.concat([df1, df2_cleaned], ignore_index=True)

# Sort by date
consolidated_df = consolidated_df.sort_values(by='Date')

# Save to a new Excel file
output_path = "consolidated_sales_v0.xlsx"
consolidated_df.to_excel(output_path, index=False)
print(f"Consolidated data saved to {output_path}")
# Display the consolidated DataFrame
print(consolidated_df.head())