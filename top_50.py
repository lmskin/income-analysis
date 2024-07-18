# This code will list the top 50 member codes for the total paid amount for all time.

import pandas as pd

# Define the file path
FILEPATH = 'All time record until 29 Jun 24.xlsx'

# Read the specific sheet from the Excel file
df = pd.read_excel(FILEPATH, sheet_name='Invoicejournal')

# Convert 'Invoice date' to datetime, handling different formats
df['Invoice date'] = pd.to_datetime(df['Invoice date'], errors='coerce')

# Drop rows with invalid dates
df = df.dropna(subset=['Invoice date'])

# Group by 'Member code' and sum the 'Paid amount'
top_members = df.groupby('Member code')['Paid amount'].sum().reset_index()

# Sort the members by total paid amount in descending order
top_members = top_members.sort_values(by='Paid amount', ascending=False).head(50)

print(top_members)

# Save the result to a CSV file
top_members.to_csv('top_50_members_paid_amount.csv', index=False)

print('done')