import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import StringIO
import re

# --- 1. CONFIGURATION AND UTILITIES ---

# Define the expected and common column names for robust parsing
# The code will search for these in the uploaded files.
EXPECTED_COLUMNS = {
    'DATE': ['Date', 'date'],
    'TIME': ['Time', 'time'],
    # Prioritize 'Total' or clear 'Active Energy'
    'ENERGY_WH': ['Total Active Energy Import(Wh)', 'Active Energy Import(Wh)', 'Energy (Wh)'],
    # Prioritize 'Total' or clear 'Active Power'
    'POWER_W': ['Total Active Power Demand(W)', 'Active Power(W)', 'Active Power(W) Avg', 'Power (W)'],
}

def find_actual_col_name(df, expected_keys):
    """Finds the actual column name in the DataFrame from a list of expected names (case-insensitive)."""
    df_cols = {col.lower(): col for col in df.columns}
    for key in expected_keys:
        if key.lower() in df_cols:
            return df_cols[key.lower()]
    return None

def preprocess_and_calculate_daily_kwh(df, msb_name):
    """
    Step 2.1: Combine Date/Time, calculate daily consumption (kWh), and prepare the load profile.
    
    Args:
        df (pd.DataFrame): The raw data for a single MSB.
        msb_name (str): The label for the MSB (e.g., 'MSB-1').

    Returns:
        tuple: (Daily consumption DataFrame, Processed load profile DataFrame)
    """
    st.info(f"Processing **{msb_name}** data...")
    
    # 1. Find Actual Column Names
    date_col = find_actual_col_name(df, EXPECTED_COLUMNS['DATE'])
    time_col = find_actual_col_name(df, EXPECTED_COLUMNS['TIME'])
    wh_col = find_actual_col_name(df, EXPECTED_COLUMNS['ENERGY_WH'])
    w_col = find_actual_col_name(df, EXPECTED_COLUMNS['POWER_W'])

    required_cols = {'Date': date_col, 'Time': time_col, 'Wh Reading': wh_col, 'Power (W)': w_col}
    
    missing_cols = [k for k, v in required_cols.items() if v is None]
    if missing_cols:
        raise ValueError(f"Missing essential columns in {msb_name}. Could not find: {', '.join(missing_cols)}.")

    # 2. Combine Date and Time into a single Timestamp
    # Combine the columns and attempt to convert to datetime.
    # Assumes common date/time formats are used.
    df['Timestamp'] = pd.to_datetime(df[date_col].astype(str) + ' ' + df[time_col].astype(str), errors='coerce', utc=True)
    
    # Drop rows where the timestamp could not be parsed
    df.dropna(subset=['Timestamp', wh_col, w_col], inplace=True)
    if df.empty:
        st.error(f"**{msb_name}**: No valid data rows after parsing timestamps and energy/power columns. Please check your file content.")
        return None, None

    # Sort and set the combined Timestamp as index for time series operations
    df.sort_values('Timestamp', inplace=True)
    df.set_index('Timestamp', inplace=True)
    
    # Clean up the energy column and convert to numeric
    df[wh_col] = pd.to_numeric(df[wh_col], errors='coerce')

    # --- Daily Consumption (Step 2.1) ---
    
    # Group by calendar day (Date)
    df_daily = df.groupby(df.index.date).agg(
        start_wh=(wh_col, 'first'),
        end_wh=(wh_col, 'last')
    )
    
    # Calculate Daily Energy Consumption (kWh)
    # Formula: Daily kWh = (End of Day Wh Reading - Start of Day Wh Reading) / 1000
    df_daily[f'{msb_name} (kWh)'] = (df_daily['end_wh'] - df_daily['start_wh']) / 1000.0
    
    # Final cleanup for the daily dataframe
    df_daily.index.names = ['Date']
    df_daily.reset_index(inplace=True)
    # Filter out any days where the calculation resulted in NaN (e.g., only one reading)
    # Also, filter out any negative values which indicate bad data (logger reset, etc.)
    df_daily = df_daily[df_daily[f'{msb_name} (kWh)'] > 0] 
    
    df_daily['Date'] = pd.to_datetime(df_daily['Date'])

    # --- Load Profile Prep (Step 3.3) ---
    
    # Prepare dataframe for load profile (instantaneous power)
    df_load_profile = df[[w_col]].copy()
    df_load_profile.rename(columns={w_col: f'{msb_name} (W)'}, inplace=True)
    
    st.success(f"**{msb_name}** processed successfully. Found {len(df_daily)} days of data.")
    return df_daily[['Date', f'{msb_name} (kWh)']], df_load_profile

def create_final_report(daily_dfs):
    """
    Step 2.2: Merge daily dataframes and calculate the Total Building Load.
    
    Args:
        daily_dfs (list): List of daily consumption DataFrames.
        
    Returns:
        pd.DataFrame: The final merged report.
    """
    if not daily_dfs:
        st.error("Cannot create a final report: no valid daily dataframes were produced.")
        return pd.DataFrame()

    # Start with the first DataFrame
    final_df = daily_dfs[0]
    
    # Merge the rest of the DataFrames
    for df in daily_dfs[1:]:
        final_df = pd.merge(final_df, df, on='Date', how='inner')
        
    # Ensure all MSB columns exist before calculating the total
    msb_cols = [col for col in final_df.columns if ' (kWh)' in col]
    if len(msb_cols) != 3:
        st.warning(f"Only {len(msb_cols)} valid MSB daily reports were merged. The total will reflect only these loads.")
    
    # Calculate Total Building Load
    final_df['Total Building Load (kWh)'] = final_df[msb_cols].sum(axis=1)
    
    # Format the Date column for the final report display
    final_df['Date'] = final_df['Date'].dt.strftime('%Y-%m-%d')
    
    # Reorder columns as specified
    col_order = ['Date'] + msb_cols + ['Total Building Load (kWh)']
    return final_df[col_order]

def calculate_average_load_profile(all_load_dfs):
    """
    Step 3.3: Calculate the average hourly power profile.
    
    Args:
        all_load_dfs (list): List of raw power DataFrames (indexed by Timestamp).
        
    Returns:
        pd.DataFrame: Average power (W) by hour of the day.
    """
    if not all_load_dfs:
        return pd.DataFrame()

    # 1. Combine all instantaneous power data
    # Use reduce and outer merge to combine all dataframes on their Timestamp index
    from functools import reduce
    df_raw_combined = reduce(lambda left, right: pd.merge(left, right, left_index=True, right_index=True, how='outer'), all_load_dfs)
    
    # 2. Calculate the Total Load
    power_cols = [col for col in df_raw_combined.columns if ' (W)' in col]
    df_raw_combined['Total Load (W)'] = df_raw_combined[power_cols].sum(axis=1)
    
    # 3. Calculate Hourly Average
    df_raw_combined['Hour'] = df_raw_combined.index.hour
    
    # Group by Hour and calculate the mean for all power columns
    df_avg_profile = df_raw_combined.groupby('Hour')[power_cols + ['Total Load (W)']].mean()
    df_avg_profile.index = df_avg_profile.index.map(lambda x: f'{x:02d}:00') # Format hour to '00:00', '01:00', etc.
    df_avg_profile.index.names = ['Hour of Day']
    df_avg_profile.reset_index(inplace=True)
    
    return df_avg_profile

# --- 2. STREAMLIT APP INTERFACE ---

st.set_page_config(
    page_title="EnergyAnalyser - Daily Energy Report Generator",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("⚡ EnergyAnalyser: Daily Energy Report Generator")
st.markdown("Upload three raw energy logger CSV files to generate a consolidated daily consumption report and load profile analysis.")

# --- File Uploader and Data Ingestion ---

uploaded_files = st.sidebar.file_uploader(
    "Upload Three MSB CSV Files",
    type="csv",
    accept_multiple_files=True
)

# Initialize variables to hold processed data
final_report_df = pd.DataFrame()
all_daily_dfs = []
all_load_profiles = []

if uploaded_files:
    if len(uploaded_files) != 3:
        st.error("🚨 **Error:** Please upload exactly **three** CSV files (MSB-1, MSB-2, MSB-3).")
    else:
        # Map MSB names to uploaded files
        msb_map = {f'MSB-{i+1}': f for i, f in enumerate(uploaded_files)}
        st.sidebar.success("3 files uploaded. Starting processing...")
        
        # --- Data Ingestion and Processing Loop ---
        for msb_name, file in msb_map.items():
            try:
                # Read the CSV file. Important: header is on the second row (index 1).
                df_raw = pd.read_csv(file, header=1)
                
                # Check for empty dataframe after reading
                if df_raw.empty:
                    st.error(f"**{msb_name}** file is empty or header is not on the second row.")
                    continue
                
                daily_df, load_profile_df = preprocess_and_calculate_daily_kwh(df_raw, msb_name)
                
                if daily_df is not None and load_profile_df is not None:
                    all_daily_dfs.append(daily_df)
                    all_load_profiles.append(load_profile_df)
                
            except Exception as e:
                st.exception(f"An error occurred while processing {msb_name}: {e}")

        # --- Final Consolidation (Step 2.2) ---
        if len(all_daily_dfs) == 3:
            final_report_df = create_final_report(all_daily_dfs)
            st.sidebar.success("✅ All data consolidated successfully.")
        elif len(all_daily_dfs) > 0:
             final_report_df = create_final_report(all_daily_dfs)
             st.sidebar.warning(f"Consolidated data from only {len(all_daily_dfs)} valid MSB files.")
        else:
            st.sidebar.error("❌ Failed to process any files. Please check the data format and column names.")
            
        # --- Calculate Average Load Profile (Step 3.3) ---
        avg_load_profile_df = calculate_average_load_profile(all_load_profiles)
        if not avg_load_profile_df.empty:
            st.sidebar.success("📈 Average Load Profile calculated.")
        

# --- 3. VISUALIZATION AND RESULTS ---

if not final_report_df.empty:
    
    # Create the tab structure for presenting results
    tab1, tab2, tab3 = st.tabs(["📊 Final Report & Download", "📈 Energy Consumption Analysis (kWh)", "⏰ Average Load Profile (W)"])
    
    # --- TAB 1: Final Data and Download (Section 3.1) ---
    with tab1:
        st.header("Final Consolidated Daily Energy Report")
        st.dataframe(final_report_df, use_container_width=True, height=400)
        
        st.subheader("Download Report")
        
        # Convert the final DataFrame to CSV for download
        csv_export = final_report_df.to_csv(index=False)
        
        st.download_button(
            label="Download Daily_Energy_Report.csv",
            data=csv_export,
            file_name='Daily_Energy_Report.csv',
            mime='text/csv',
            help='Click to download the combined and processed daily energy consumption data.'
        )

    # --- TAB 2: Energy Consumption Analysis (Section 3.2) ---
    with tab2:
        st.header("Daily Energy Consumption Analysis")
        
        # Chart 1: Total Daily Energy Consumption (Bar Chart)
        st.subheader("Total Daily Energy Consumption (kWh)")
        
        fig_total = px.bar(
            final_report_df, 
            x='Date', 
            y='Total Building Load (kWh)',
            title="Total Daily Energy Consumption (kWh)",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_total.update_layout(xaxis_title="Date", yaxis_title="Total Building Load (kWh)")
        st.plotly_chart(fig_total, use_container_width=True)
        
        st.markdown("---")
        
        # Chart 2: Load Apportionment Over Time (Stacked Area Chart)
        st.subheader("MSB Load Apportionment Over Time")
        
        # Melt the dataframe to have one row per Date and MSB for Plotly
        df_melted = final_report_df.melt(
            id_vars=['Date', 'Total Building Load (kWh)'], 
            value_vars=[col for col in final_report_df.columns if 'MSB-' in col],
            var_name='MSB', 
            value_name='Consumption (kWh)'
        )

        fig_stacked = px.area(
            df_melted, 
            x='Date', 
            y='Consumption (kWh)', 
            color='MSB',
            title="MSB Load Apportionment Over Time",
            color_discrete_sequence=px.colors.qualitative.Dark24,
            markers=True # Adds dots to highlight data points
        )
        fig_stacked.update_layout(
            xaxis_title="Date", 
            yaxis_title="Energy Consumption (kWh)",
            hovermode="x unified"
        )
        st.plotly_chart(fig_stacked, use_container_width=True)

    # --- TAB 3: Average Load Profile (Section 3.3) ---
    with tab3:
        st.header("24-Hour Average Load Profile (W)")
        
        if not avg_load_profile_df.empty:
            
            # Melt the dataframe for Plotly to handle multiple lines
            df_profile_melted = avg_load_profile_df.melt(
                id_vars='Hour of Day',
                var_name='Load Source',
                value_name='Average Active Power (W)'
            )
            
            # Use Plotly Express for the line chart
            fig_profile = px.line(
                df_profile_melted,
                x='Hour of Day',
                y='Average Active Power (W)',
                color='Load Source',
                title="24-Hour Average Total Load Profile (W)",
                markers=True, # Add markers for clarity
                color_discrete_map={'Total Load (W)': 'black'} # Highlight total load
            )

            # Improve x-axis for better time reading
            fig_profile.update_layout(
                xaxis_title="Hour of Day",
                yaxis_title="Average Active Power (W)",
                hovermode="x unified",
                legend_title="Load Source"
            )
            
            # Ensure the line chart is plotted
            st.plotly_chart(fig_profile, use_container_width=True)
            
            st.subheader("Raw Average Data")
            st.dataframe(avg_load_profile_df, use_container_width=True)
            
        else:
            st.warning("Could not calculate the Average Load Profile due to insufficient or invalid raw power data.")

else:
    st.info("⬆️ Please upload your three MSB data files in the sidebar to begin the analysis.")
