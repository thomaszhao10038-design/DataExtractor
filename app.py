import streamlit as st
import pandas as pd
import io

# --- Configuration ---
# Assuming the relevant Active Power (W) column for aggregation is the 142nd column
# in the CSV when read without skipping rows (Index 141 in a 0-based system).
# This index is based on the structure of the uploaded 'raw data MSB X.csv' files.
ACTIVE_POWER_COL_INDEX = 141

# Headers for the Final Consolidated Report
REPORT_COLUMNS = ['Date', 'MSB-1 (kW)', 'MSB-2 (kW)', 'MSB-3 (kW)', 'Total Building Load (kW)']

# --- Helper Functions ---

@st.cache_data
def load_and_process_raw_data(uploaded_file, name):
    """
    Reads the raw MSB CSV data, extracts the daily total energy usage (in kWh, 
    which is the sum of Active Power (W) readings), and converts it to average daily power (kW).
    
    The raw data CSV is unusual: the actual data starts on the 3rd row (index 2), 
    but the relevant column index (141) is determined from the full 0-based file structure.
    """
    if uploaded_file is None:
        return None
    
    try:
        # 1. Read the full CSV without skipping the header, using the detected column index
        # We read the whole file content into a string buffer first to handle large files efficiently
        uploaded_file.seek(0)
        file_content = uploaded_file.read().decode('utf-8')
        
        # We need to manually specify dtypes as columns 141+ are often mixed type in the raw input
        # We load only the timestamp column (index 1) and the target data column (index 141)
        # We skip the first row (ProductSN)
        df = pd.read_csv(
            io.StringIO(file_content),
            header=2, # The column names we care about are on the 3rd row (index 2: Date, Time, ...)
            usecols=[1, ACTIVE_POWER_COL_INDEX],
            dtype={1: str, ACTIVE_POWER_COL_INDEX: str}, # Treat as string initially to handle potential bad data
            skiprows=[1], # Skip the second row which contains the long list of units and productSN
            encoding='utf-8',
        )

        # Rename columns for clarity
        df.columns = ['DateTime', 'Active Power (W)']
        
        # Clean data: Remove rows where Active Power (W) is NaN or non-numeric
        df['Active Power (W)'] = pd.to_numeric(df['Active Power (W)'], errors='coerce')
        df = df.dropna(subset=['Active Power (W)'])
        
        if df.empty:
            st.warning(f"No valid power data found in {name} file after cleaning.")
            return None
            
        # Extract the date part and convert to datetime objects
        # The DateTime column looks like "2025-11-11,00:00:00"
        # We combine this with the 'Date' column from the actual header row to form the full timestamp, 
        # but since pandas read it as one column, we need to split it.
        
        df['Date'] = pd.to_datetime(df['DateTime'].apply(lambda x: x.split(',')[0]), errors='coerce')

        # Calculate daily sum and average for the raw Active Power (W) readings
        # The sum of 15-minute Active Power (W) readings represents (Total Watt-Minutes / 4) in Watt
        # To get the average kW: (Sum of W / Count of Readings) / 1000
        
        # Calculate the count of readings per day (assuming 96 readings per full day (4 per hour * 24 hours))
        daily_counts = df.groupby(df['Date'].dt.date)['Active Power (W)'].count()
        
        # Calculate the sum of power readings per day (Watt-Readings)
        daily_sum_W = df.groupby(df['Date'].dt.date)['Active Power (W)'].sum()
        
        # Calculate the average Active Power in W for the day: Sum(W) / Count(Readings)
        daily_avg_W = daily_sum_W / daily_counts
        
        # Convert daily average Power (W) to kW (kW = W / 1000)
        daily_avg_kW = daily_avg_W / 1000
        
        # Convert the index (date objects) back to a DataFrame
        result_df = daily_avg_kW.reset_index()
        result_df.columns = ['Date', name]
        result_df['Date'] = pd.to_datetime(result_df['Date'])
        
        return result_df

    except Exception as e:
        st.error(f"Error processing {name} data: {e}")
        return None

def merge_and_finalize_report(df_msb1, df_msb2, df_msb3):
    """Merges the three MSB dataframes and calculates the total load."""
    
    # Start with MSB-1
    final_df = df_msb1.set_index('Date')

    # Merge MSB-2 and MSB-3
    final_df = final_df.join(df_msb2.set_index('Date'), how='outer')
    final_df = final_df.join(df_msb3.set_index('Date'), how='outer')

    # Fill NaN values (for days where one MSB has no data) with 0 for calculation
    # Only fill after merging, so we can calculate the total correctly
    final_df = final_df.fillna(0)
    
    # Calculate Total Building Load
    final_df['Total Building Load (kW)'] = final_df.sum(axis=1)

    # Clean up and reset index
    final_df = final_df.reset_index()
    final_df.columns = REPORT_COLUMNS
    
    # Sort by date
    final_df = final_df.sort_values(by='Date')
    
    # Format Date column for final display (YYYY-MM-DD)
    final_df['Date'] = final_df['Date'].dt.strftime('%Y-%m-%d')
    
    return final_df

# --- Streamlit UI ---

st.set_page_config(
    page_title="EnergyAnalyser: MSB Daily Load Report Generator",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("⚡ MSB Daily Load Report Generator")
st.markdown("Upload the raw CSV data files for each Main Switch Board (MSB) to generate a consolidated daily load report.")

# File Uploader in the sidebar
with st.sidebar:
    st.header("1. Upload Raw Data Files")
    msb1_file = st.file_uploader("Upload raw data MSB-1 CSV", type=['csv'], key="msb1")
    msb2_file = st.file_uploader("Upload raw data MSB-2 CSV", type=['csv'], key="msb2")
    msb3_file = st.file_uploader("Upload raw data MSB-3 CSV", type=['csv'], key="msb3")
    
    st.markdown("---")
    st.info(
        f"**Technical Note:** This app assumes the 'Total Active Power (W)' data is located at **Column Index {ACTIVE_POWER_COL_INDEX}** (142nd column) in your raw CSV files."
    )

if msb1_file and msb2_file and msb3_file:
    
    st.header("2. Processing Data")

    # Load and process data
    df_msb1 = load_and_process_raw_data(msb1_file, 'MSB-1 (kW)')
    df_msb2 = load_and_process_raw_data(msb2_file, 'MSB-2 (kW)')
    df_msb3 = load_and_process_raw_data(msb3_file, 'MSB-3 (kW)')
    
    # Check if all data frames were successfully processed
    if df_msb1 is not None and df_msb2 is not None and df_msb3 is not None:
        
        st.success("Successfully processed raw data from all three MSBs.")
        
        # Merge and finalize the report
        final_report_df = merge_and_finalize_report(df_msb1, df_msb2, df_msb3)
        
        st.header("3. Consolidated Daily Load Report")
        
        # Display the final table
        st.dataframe(
            final_report_df, 
            use_container_width=True,
            hide_index=True,
            column_config={
                "Date": st.column_config.DateColumn("Date", format="YYYY-MM-DD"),
                "MSB-1 (kW)": st.column_config.NumberColumn("MSB-1 (kW)", format="%.3f"),
                "MSB-2 (kW)": st.column_config.NumberColumn("MSB-2 (kW)", format="%.3f"),
                "MSB-3 (kW)": st.column_config.NumberColumn("MSB-3 (kW)", format="%.3f"),
                "Total Building Load (kW)": st.column_config.NumberColumn("Total Building Load (kW)", format="%.3f"),
            }
        )
        
        # Download button
        csv_export = final_report_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Consolidated Report as CSV",
            data=csv_export,
            file_name='Consolidated_Daily_Load_Report.csv',
            mime='text/csv',
            key='download-csv'
        )
        
        # Optional: Display a simple trend chart
        st.subheader("Daily Total Load Trend")
        chart_df = final_report_df[['Date', 'Total Building Load (kW)']].rename(columns={'Total Building Load (kW)': 'Total Load'})
        chart_df['Date'] = pd.to_datetime(chart_df['Date']) # Convert back to datetime for chart
        st.line_chart(chart_df, x='Date', y='Total Load', use_container_width=True)


else:
    st.info("Please upload all three MSB raw data files to generate the report.")

st.sidebar.markdown("---")
st.sidebar.caption("EnergyAnalyser v1.0")
