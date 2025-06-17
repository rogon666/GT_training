###############################################################################
# Loading Excel files with pandas
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