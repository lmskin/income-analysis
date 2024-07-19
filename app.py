import pandas as pd
import plotly.express as px
import streamlit as st

# Streamlit app title
st.title("Excel Data Analysis")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

# Analysis type selection
analysis_type = st.selectbox("Choose analysis type:", ["Monthly Paid Amount", "Quarterly Paid Amount", "Yearly Paid Amount", "Top 50 Members"])

if uploaded_file is not None:
    # Read the specific sheet from the uploaded Excel file
    df = pd.read_excel(uploaded_file, sheet_name='Invoicejournal')

    # Convert 'Invoice date' to datetime, handling different formats
    df['Invoice date'] = pd.to_datetime(df['Invoice date'], errors='coerce')

    # Drop rows with invalid dates
    df = df.dropna(subset=['Invoice date'])

    if analysis_type in ["Monthly Paid Amount", "Quarterly Paid Amount", "Yearly Paid Amount"]:
        if analysis_type == "Monthly Paid Amount":
            # Extract year and month from 'Invoice date'
            df['Period'] = df['Invoice date'].dt.to_period('M')
            period_format = 'M'
        elif analysis_type == "Quarterly Paid Amount":
            # Extract year and quarter from 'Invoice date'
            df['Period'] = df['Invoice date'].dt.to_period('Q')
            period_format = 'Q'
        elif analysis_type == "Yearly Paid Amount":
            # Extract year from 'Invoice date'
            df['Period'] = df['Invoice date'].dt.to_period('Y')
            period_format = 'Y'

        # Group by 'Period' and sum the 'Paid amount'
        paid_amount = df.groupby('Period')['Paid amount'].sum().reset_index()

        # Convert Period to string for Plotly
        paid_amount['Period'] = paid_amount['Period'].astype(str)

        # Create an interactive bar chart using Plotly
        fig = px.bar(
            paid_amount,
            x='Period',
            y='Paid amount',
            title=f'{analysis_type} for All Time',
            labels={'Period': 'Period', 'Paid amount': 'Paid Amount'},
            hover_data={'Paid amount': ':.2f'}  # Format hover data to 2 decimal places
        )

        # Show the plot in the Streamlit app
        st.plotly_chart(fig)

    elif analysis_type == "Top 50 Members":
        # Group by 'Member code' and sum the 'Paid amount'
        top_members = df.groupby(['Member code', 'Member name'])['Paid amount'].sum().reset_index()

        # Sort the members by total paid amount in descending order
        top_members = top_members.sort_values(by='Paid amount', ascending=False).head(50)

        # Display the top 50 members in the Streamlit app
        st.subheader("Top 50 Members by Total Paid Amount")
        st.dataframe(top_members)

        # Save the result to a CSV file
        top_members.to_csv('top_50_members_paid_amount.csv', index=False)
        st.success('Top 50 members saved to CSV file.')

# To run the Streamlit app, save this file as app.py and run `streamlit run app.py` in your terminal.
