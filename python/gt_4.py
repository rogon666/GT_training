###############################################################################
# Anomaly Detection in Sales Data using Isolation Forest and Elliptic Envelopes
# ----------------------------------------------------------------------------
# This script detects anomalies in sales data using Isolation Forest and 
# Elliptic Envelope methods.
###############################################################################
import os
import pandas as pd

new_dir = "C:/2025/06_june/GT/excel"
os.chdir(new_dir)
print("New Directory:", os.getcwd())

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
model = IsolationForest(contamination=0.1, random_state=42)  # 10% contamination
model.fit(X_scaled)

# Predict anomalies (1 = normal, -1 = anomaly)
df['anomaly'] = model.predict(X_scaled)

# Add anomaly scores (the lower the score, the more anomalous)
df['anomaly_score'] = model.decision_function(X_scaled)

# Identify anomalies
anomalies = df[df['anomaly'] == -1]
print("\nDetected anomalies (Isolation Forest):")
print(anomalies[['Date', 'Product', 'Units Sold', 'Unit Price ($)', 'Total Sales ($)', 'anomaly_score']])

# Visualize the results for Isolation Forest
plt.figure(figsize=(12, 6))
plt.scatter(df['Date'], df['Total Sales ($)'], 
            c=df['anomaly'], cmap='coolwarm', 
            label='Normal')
plt.scatter(anomalies['Date'], anomalies['Total Sales ($)'], 
            color='red', marker='x', s=100, 
            label='Anomaly')
plt.title('Total Sales with Anomalies Highlighted (Isolation Forest)')
plt.xlabel('Date')
plt.ylabel('Total Sales ($)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

#-----------------------------------------------
# Anomaly Detection in Sales Data using Elliptic Envelope
from sklearn.covariance import EllipticEnvelope

# Train the Elliptic Envelope model
envelope = EllipticEnvelope(contamination=0.1, random_state=42)
envelope.fit(X_scaled)

# Predict anomalies (1 = normal, -1 = anomaly)
df['elliptic_anomaly'] = envelope.predict(X_scaled)

# Add anomaly scores (the lower, the more anomalous)
df['elliptic_anomaly_score'] = envelope.decision_function(X_scaled)

# Identify anomalies
elliptic_anomalies = df[df['elliptic_anomaly'] == -1]
print("\nDetected anomalies (Elliptic Envelope):")
print(elliptic_anomalies[['Date', 'Product', 'Units Sold', 'Unit Price ($)', 'Total Sales ($)', 'elliptic_anomaly_score']])

# Visualize the results for Elliptic Envelope
plt.figure(figsize=(12, 6))
plt.scatter(df['Date'], df['Total Sales ($)'], 
            c=df['elliptic_anomaly_score'], cmap='coolwarm', 
            label='Normal')
plt.scatter(elliptic_anomalies['Date'], elliptic_anomalies['Total Sales ($)'], 
            color='red', marker='x', s=100, 
            label='Anomaly')
plt.title('Total Sales with Anomalies Highlighted (Elliptic Envelope)')
plt.xlabel('Date')
plt.ylabel('Total Sales ($)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# Save results to a new Excel file
df.to_excel('sales_with_IFEE_anomalies.xlsx', index=False)
print("\nResults saved to 'sales_with_IFEE_anomalies.xlsx'")