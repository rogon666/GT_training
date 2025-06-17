###############################################################################
# Cleaning and Consolidating Sales Data from Excel Files
# This script consolidates sales data from two Excel files, cleans the data, 
# and saves the final consolidated result to a new Excel file.
###############################################################################
import os
import pandas as pd
from word2number import w2n

# Check current directory
print("Current Directory:", os.getcwd())
# Change to a new directory (example path)
new_dir = "C:/2025/06_june/GT/excel"
os.chdir(new_dir)
# Confirm change
print("New Directory:", os.getcwd())

# Load the Excel files
df1 = pd.read_excel("monthly_sales_v3.xlsx", sheet_name="Sheet1")  # or sheet_name=None for all sheets
df2 = pd.read_excel("monthly_sales_v4 - Copy final use this.xlsx", sheet_name="Sheet1")
print("DataFrame 1:")
print(df1.head())  
print("DataFrame 2:")
print(df2.head()) 

# -----------------------------------------------------------------------------
# Fix for text values in 'Units Sold' for df1
text_to_number = {
    'TWO': 2,
    'two': 2
}

# Replace text entries in 'Units Sold' for df1
df1['Units Sold'] = df1['Units Sold'].replace(text_to_number)

# Function to convert textual numbers to integers
def convert_text_to_number(val):
    try:
        if isinstance(val, str):
            return w2n.word_to_num(val.lower())
        return val
    except:
        return val  # Return the original value if conversion fails

# Rename columns in df2 to match df1
df2_cleaned = df2.rename(columns={
    'code': 'Product',
    'units': 'Units Sold',
    'Price ($)': 'Unit Price ($)',
    'Total ($)': 'Total Sales ($)'
})

# Apply the conversion function to 'Units Sold'
df2_cleaned['Units Sold'] = df2_cleaned['Units Sold'].apply(convert_text_to_number)

# Clean 'Total Sales ($)' by removing commas (if any) and convert to float
df2_cleaned['Total Sales ($)'] = df2_cleaned['Total Sales ($)'].replace({',': ''}, regex=True).astype(float)

# Ensure numeric conversion for both Units Sold and Unit Price
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
output_path = "consolidated_sales.xlsx"
consolidated_df.to_excel(output_path, index=False)

output_path
