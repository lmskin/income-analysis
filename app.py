import pandas as pd
import plotly.express as px
import streamlit as st

# Streamlit app title
st.title("Excel Data Analysis")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

# Analysis type selection
analysis_type = st.selectbox("Choose analysis type:", ["Monthly Paid Amount", "Quarterly Paid Amount", "Yearly Paid Amount", "Top 50 Members", "Compare Month Over Years"])

# Month selection for comparison (only show when "Compare Month Over Years" is selected)
month_names = ["January", "February", "March", "April", "May", "June", 
               "July", "August", "September", "October", "November", "December"]
selected_month = None
if analysis_type == "Compare Month Over Years":
    selected_month = st.selectbox("Select a month to compare across years:", month_names)

if uploaded_file is not None:
    # Read the specific sheet from the uploaded Excel file (skip first 10 metadata rows)
    df = pd.read_excel(uploaded_file, sheet_name='Invoicejournal', header=10)

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

    elif analysis_type == "Compare Month Over Years":
        # Get the month number from the selected month name
        month_number = month_names.index(selected_month) + 1
        
        # Extract year and month from 'Invoice date'
        df['Year'] = df['Invoice date'].dt.year
        df['Month'] = df['Invoice date'].dt.month
        
        # Filter data for the selected month
        month_data = df[df['Month'] == month_number]
        
        if month_data.empty:
            st.warning(f"No data available for {selected_month}.")
        else:
            # Group by year and sum the 'Paid amount'
            yearly_comparison = month_data.groupby('Year')['Paid amount'].sum().reset_index()
            
            # Convert Year to string for better display
            yearly_comparison['Year'] = yearly_comparison['Year'].astype(str)
            
            # Create an interactive bar chart using Plotly
            fig = px.bar(
                yearly_comparison,
                x='Year',
                y='Paid amount',
                title=f'{selected_month} Sales Comparison Across Years',
                labels={'Year': 'Year', 'Paid amount': 'Paid Amount'},
                hover_data={'Paid amount': ':.2f'},
                color='Paid amount',
                color_continuous_scale='Blues'
            )
            
            # Customize layout
            fig.update_layout(
                xaxis_title="Year",
                yaxis_title="Paid Amount",
                showlegend=False
            )
            
            # Show the plot in the Streamlit app
            st.plotly_chart(fig)
            
            # Display summary statistics
            st.subheader(f"Summary for {selected_month}")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total", f"${yearly_comparison['Paid amount'].sum():,.2f}")
            with col2:
                st.metric("Average", f"${yearly_comparison['Paid amount'].mean():,.2f}")
            with col3:
                max_year = yearly_comparison.loc[yearly_comparison['Paid amount'].idxmax(), 'Year']
                st.metric("Best Year", max_year)
            
            # Display the data table
            st.dataframe(yearly_comparison.rename(columns={'Paid amount': 'Paid Amount'}))

# To run the Streamlit app, save this file as app.py and run `streamlit run app.py` in your terminal.
