###############################################################################
# Anomaly Detection in Sales Data using Isolation Forest 
# ----------------------------------------------------------------------------
# This script detects anomalies in sales data using the IF algorithm
# #############################################################################
import os
import pandas as pd

new_dir = "C:/2025/06_june/GT/excel"
os.chdir(new_dir)
print("New Directory:", os.getcwd())

# Load the consolidated sales data
df = pd.read_excel("consolidated_sales.xlsx", sheet_name="Sheet1") 
print(df.head())  

#-------------------------------------------------------------------------------
# Anomaly Detection in Sales Data using Isolation Forest
import numpy as np
import matplotlib.pyplot as plt
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler

# Prepare the data for anomaly detection
X = df[['Total Sales ($)']].values

# Standardize the data
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# Train the Isolation Forest model
model = IsolationForest(contamination=0.1, random_state=42)  # 5% contamination
model.fit(X_scaled)

# Predict anomalies (1 = normal, -1 = anomaly)
df['anomaly'] = model.predict(X_scaled)

# Add anomaly scores (the lower the score, the more anomalous)
df['anomaly_score'] = model.decision_function(X_scaled)

# Identify anomalies
anomalies = df[df['anomaly'] == -1]
print("\nDetected anomalies:")
print(anomalies[['Date', 'Product', 'Units Sold', 'Unit Price ($)', 'Total Sales ($)', 'anomaly_score']])

# Visualize the results
plt.figure(figsize=(12, 6))

# Plot all sales data
plt.scatter(df['Date'], df['Total Sales ($)'], 
            c=df['anomaly'], cmap='coolwarm', 
            label='Normal')

# Highlight anomalies
plt.scatter(anomalies['Date'], anomalies['Total Sales ($)'], 
            color='red', marker='x', s=100, 
            label='Anomaly')

plt.title('Total Sales with Anomalies Highlighted')
plt.xlabel('Date')
plt.ylabel('Total Sales ($)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# Save results to a new Excel file
df.to_excel('sales_with_IF_anomalies.xlsx', index=False)
print("\nResults saved to 'sales_with_IF_anomalies.xlsx'")