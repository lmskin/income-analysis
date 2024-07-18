# This code will read the Excel file, process the data, and create an interactive bar chart for the monthly paid amount.

import pandas as pd
import plotly.express as px

# Define the file path
FILEPATH = "C:\\Users\\user\\Desktop\\Python\\income-analysis\\All time record until 29 Jun 24.xlsx"

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

# Convert YearMonth to string for Plotly
monthly_paid_amount['YearMonth'] = monthly_paid_amount['YearMonth'].astype(str)

# Create an interactive bar chart using Plotly
fig = px.bar(
    monthly_paid_amount,
    x='YearMonth',
    y='Paid amount',
    title='Monthly Paid Amount for All Time',
    labels={'YearMonth': 'Month', 'Paid amount': 'Paid Amount'},
    hover_data={'Paid amount': ':.2f'}  # Format hover data to 2 decimal places
)

# Show the plot
fig.show()

print('done')
