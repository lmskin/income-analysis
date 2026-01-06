import pandas as pd
import streamlit as st
import plotly.express as px

# Function to read all sheets from an uploaded Excel file
def read_all_sheets_from_excel(uploaded_file):
    # Read all sheets into a dictionary of dataframes
    return pd.read_excel(uploaded_file, sheet_name=None)

# Streamlit app
st.title('Interactive Treatment Value Analysis')

# Upload the Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the data from the uploaded Excel file
    dataframes = read_all_sheets_from_excel(uploaded_file)
    
    # Extract the relevant dataframe
    df = dataframes['trt_journal_report']
    
    # Convert 'Treatment date' to datetime format
    df['Treatment date'] = pd.to_datetime(df['Treatment date'], format='%m-%d-%y')
    
    # Group by month and year, then sum the 'Treatment value'
    monthly_treatment_value = df.groupby(df['Treatment date'].dt.to_period('M'))['Treatment value'].sum().reset_index()
    monthly_treatment_value['Treatment date'] = monthly_treatment_value['Treatment date'].astype(str)  # Convert period to string for plotly
    
    # Show the dataframe
    st.write("Monthly Treatment Value Data:")
    st.write(monthly_treatment_value)
    
    # Plot the interactive bar chart using Plotly
    fig = px.bar(monthly_treatment_value, 
                 x='Treatment date', 
                 y='Treatment value', 
                 title='Total Monthly Treatment Value Over Time',
                 labels={'Treatment date': 'Month-Year', 'Treatment value': 'Total Treatment Value'})
    
    fig.update_layout(xaxis_tickangle=-45)  # Rotate x-axis labels
    
    # Display the interactive plot in Streamlit
    st.plotly_chart(fig)
