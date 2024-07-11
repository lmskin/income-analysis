# This code will read the Excel file, process the data, and create a bar chart for the monthly paid amount.

import pandas as pd
import matplotlib.pyplot as plt

# Define the file path
FILEPATH = 'All time record until 29 Jun 24.xlsx'

# Read the specific sheet from the Excel file
df = pd.read_excel(FILEPATH, sheet_name='Invoicejournal')

# Convert 'Invoice date' to datetime, handling different formats
df['Invoice date'] = pd.to_datetime(df['Invoice date'], errors='coerce')

# Drop rows with invalid dates
df = df.dropna(subset=['Invoice date'])

# Extract year and month from 'Invoice date'
df['YearMonth'] = df['Invoice date'].dt.to_period('M')

# Group by 'YearMonth' and sum the 'Paid amount'
monthly_paid_amount = df.groupby('YearMonth')['Paid amount'].sum().reset_index()

# Plot the bar chart
plt.figure(figsize=(12, 6))
plt.bar(monthly_paid_amount['YearMonth'].astype(str), monthly_paid_amount['Paid amount'])
plt.xlabel('Month')
plt.ylabel('Paid Amount')
plt.title('Monthly Paid Amount for All Time')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()

print('done')