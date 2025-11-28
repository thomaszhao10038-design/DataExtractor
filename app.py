import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from functools import reduce

# --- 1. CONFIGURATION AND UTILITIES ---

# Define a broader set of expected and common column names
# **Crucially, if the script fails to find a column, check your CSV header and add the EXACT name here.**
EXPECTED_COLUMNS = {
    # Add your specific logger headers to these lists if default ones fail (e.g., 'START_DATE')
    'DATE': ['Date', 'date', 'Recording Date', 'Day', 'Log Date'],
    'TIME': ['Time', 'time', 'Recording Time', 'Hour', 'Log Time'],
    # Energy (Wh) - Look for total energy import
    'ENERGY_WH': ['Total Active Energy Import(Wh)', 'Active Energy Import(Wh)', 'Energy (Wh)', 'Wh Reading', 'Total Energy'],
    # Power (W) - Look for average or instantaneous active power
    'POWER_W': ['Total Active Power Demand(W)', 'Active Power(W)', 'Active Power(W) Avg', 'Power (W)', 'Total Power'],
}

def find_actual_col_name(df, expected_keys):
    """
    Finds the actual column name in the DataFrame from a list of expected names (case-insensitive and trimmed).
    """
    # 1. Normalize DataFrame columns (trim whitespace and lower case)
    df_cols_map = {col.strip().lower(): col for col in df.columns}
    
    # 2. Search for the expected keys
    for key in expected_keys:
        normalized_key = key.strip().lower()
        if normalized_key in df_cols_map:
            return df_cols_map[normalized_key]
    
    return None

def preprocess_and_calculate_daily_kwh(df, msb_name):
    """
    Combines Date/Time, calculates daily consumption (kWh), and prepares the load profile.
    
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

    required_cols_map = {'Date': date_col, 'Time': time_col, 'Energy (Wh)': wh_col, 'Power (W)': w_col}
    
    missing_cols = [k for k, v in required_cols_map.items() if v is None]
    if missing_cols:
        # Generate a list of all potential names that were searched for but not found
        missing_search_terms = []
        if 'Date' in missing_cols: missing_search_terms.extend(EXPECTED_COLUMNS['DATE'])
        if 'Time' in missing_cols: missing_search_terms.extend(EXPECTED_COLUMNS['TIME'])
        
        st.error(f"❌ **Error in {msb_name}:** Missing essential columns: {', '.join(missing_cols)}.")
        st.warning(f"**Tip:** The script searched for names like: {', '.join(list(set(missing_search_terms))[:5])}... (and others). Please add the **exact header names** from your CSV file to the `EXPECTED_COLUMNS` dictionary in `app.py`.")
        return None, None

    # 2. Combine Date and Time into a single Timestamp
    try:
        # Combine the columns and attempt to convert to datetime.
        df['Timestamp'] = pd.to_datetime(df[date_col].astype(str).str.strip() + ' ' + df[time_col].astype(str).str.strip(), 
                                         errors='coerce', utc=True)
    except Exception as e:
        st.error(f"❌ **Error in {msb_name}:** Failed to combine '{date_col}' and '{time_col}' into a datetime object. Check data format.")
        return None, None
        
    # Drop rows where the timestamp could not be parsed or energy/power is NaN
    df.dropna(subset=['Timestamp', wh_col, w_col], inplace=True)
    
    if df.empty:
        st.error(f"**{msb_name}**: No valid data rows after cleaning. Check for empty or non-numeric readings.")
        return None, None

    # Sort and set the combined Timestamp as index for time series operations
    df.sort_values('Timestamp', inplace=True)
    df.set_index('Timestamp', inplace=True)
    
    # Clean up the energy column and convert to numeric
    df[wh_col] = pd.to_numeric(df[wh_col], errors='coerce')
    df[w_col] = pd.to_numeric(df[w_col], errors='coerce')

    # --- Daily Consumption Calculation ---
    
    # Group by calendar day (Date)
    df_daily = df.groupby(df.index.date).agg(
        start_wh=(wh_col, 'first'),
        end_wh=(wh_col, 'last')
    )
    
    # Calculate Daily Energy Consumption (kWh)
    daily_consumption_col = f'{msb_name} (kWh)'
    df_daily[daily_consumption_col] = (df_daily['end_wh'] - df_daily['start_wh']) / 1000.0
    
    # Filter out negative/zero consumption (due to bad data or meter reset)
    df_daily = df_daily[df_daily[daily_consumption_col] > 0] 
    
    df_daily.index.names = ['Date']
    df_daily.reset_index(inplace=True)
    df_daily['Date'] = pd.to_datetime(df_daily['Date'])

    # --- Load Profile Prep ---
    
    df_load_profile = df[[w_col]].copy()
    df_load_profile.rename(columns={w_col: f'{msb_name} (W)'}, inplace=True)
    
    st.success(f"✅ **{msb_name}** processed successfully. Found {len(df_daily)} days of data.")
    return df_daily[['Date', daily_consumption_col]], df_load_profile

def create_final_report(daily_dfs):
    """Merges daily dataframes and calculates the Total Building Load."""
    if not daily_dfs:
        return pd.DataFrame()

    # Merge all daily dataframes on the common 'Date' column
    final_df = reduce(lambda left, right: pd.merge(left, right, on='Date', how='inner'), daily_dfs)
    
    # Identify MSB columns
    msb_cols = [col for col in final_df.columns if ' (kWh)' in col]
    
    # Calculate Total Building Load
    final_df['Total Building Load (kWh)'] = final_df[msb_cols].sum(axis=1)
    
    final_df['Date'] = final_df['Date'].dt.strftime('%Y-%m-%d')
    
    # Reorder columns
    col_order = ['Date'] + msb_cols + ['Total Building Load (kWh)']
    return final_df[col_order]

def calculate_average_load_profile(all_load_dfs):
    """Calculates the average hourly power profile."""
    if not all_load_dfs:
        return pd.DataFrame()

    # Combine all instantaneous power data
    df_raw_combined = reduce(lambda left, right: pd.merge(left, right, left_index=True, right_index=True, how='outer'), all_load_dfs)
    
    # Calculate the Total Load
    power_cols = [col for col in df_raw_combined.columns if ' (W)' in col and 'Total' not in col]
    df_raw_combined['Total Load (W)'] = df_raw_combined[power_cols].sum(axis=1)
    
    # Calculate Hourly Average
    df_raw_combined['Hour'] = df_raw_combined.index.hour
    
    # Group by Hour and calculate the mean for all power columns
    all_cols = power_cols + ['Total Load (W)']
    df_avg_profile = df_raw_combined.groupby('Hour')[all_cols].mean()
    
    # Format index and reset
    df_avg_profile.index = df_avg_profile.index.map(lambda x: f'{x:02d}:00') 
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
    "Upload Three MSB CSV Files (MSB-1, MSB-2, MSB-3)",
    type="csv",
    accept_multiple_files=True
)

# Initialize variables to hold processed data
final_report_df = pd.DataFrame()
all_daily_dfs = []
all_load_profiles = []

if uploaded_files:
    if len(uploaded_files) != 3:
        st.error("🚨 **Error:** Please upload exactly **three** CSV files (MSB-1, MSB-2, MSB-3) for a complete analysis.")
    else:
        msb_map = {f'MSB-{i+1}': f for i, f in enumerate(uploaded_files)}
        st.sidebar.success("3 files uploaded. Starting processing...")
        
        # --- Data Ingestion and Processing Loop ---
        for msb_name, file in msb_map.items():
            try:
                # Read the CSV file. Crucially, header is on the second row (index 1).
                df_raw = pd.read_csv(file, header=1, skipinitialspace=True)
                
                if df_raw.empty:
                    st.error(f"**{msb_name}** file is empty or the header on row 2 is incorrect.")
                    continue
                
                daily_df, load_profile_df = preprocess_and_calculate_daily_kwh(df_raw, msb_name)
                
                if daily_df is not None and load_profile_df is not None:
                    all_daily_dfs.append(daily_df)
                    all_load_profiles.append(load_profile_df)
                
            except Exception as e:
                st.exception(f"An unexpected error occurred while processing {msb_name}. This is usually a file formatting issue: {e}")

        # --- Final Consolidation ---
        if len(all_daily_dfs) >= 1:
            final_report_df = create_final_report(all_daily_dfs)
            st.sidebar.success(f"✅ Daily Reports consolidated from {len(all_daily_dfs)} MSB files.")
        else:
            st.sidebar.error("❌ Failed to process any files successfully. Check error messages above.")
            
        # --- Calculate Average Load Profile ---
        avg_load_profile_df = pd.DataFrame()
        if all_load_profiles:
            avg_load_profile_df = calculate_average_load_profile(all_load_profiles)
            st.sidebar.success("📈 Average Load Profile calculated.")
        

# --- 3. VISUALIZATION AND RESULTS ---

if not final_report_df.empty:
    
    tab1, tab2, tab3 = st.tabs(["📊 Final Report & Download", "📈 Energy Consumption Analysis (kWh)", "⏰ Average Load Profile (W)"])
    
    # --- TAB 1: Final Data and Download ---
    with tab1:
        st.header("Final Consolidated Daily Energy Report")
        st.dataframe(final_report_df, use_container_width=True, height=400)
        
        st.subheader("Download Report")
        
        csv_export = final_report_df.to_csv(index=False).encode('utf-8')
        
        st.download_button(
            label="Download Daily_Energy_Report.csv",
            data=csv_export,
            file_name='Daily_Energy_Report.csv',
            mime='text/csv',
            help='Click to download the combined and processed daily energy consumption data.'
        )

    # --- TAB 2: Energy Consumption Analysis ---
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
        
        msb_cols = [col for col in final_report_df.columns if 'MSB-' in col]
        df_melted = final_report_df[['Date'] + msb_cols].melt(
            id_vars='Date', 
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
        )
        fig_stacked.update_layout(
            xaxis_title="Date", 
            yaxis_title="Energy Consumption (kWh)",
            hovermode="x unified"
        )
        st.plotly_chart(fig_stacked, use_container_width=True)

    # --- TAB 3: Average Load Profile ---
    with tab3:
        st.header("24-Hour Average Load Profile (W)")
        
        if not avg_load_profile_df.empty:
            
            df_profile_melted = avg_load_profile_df.melt(
                id_vars='Hour of Day',
                var_name='Load Source',
                value_name='Average Active Power (W)'
            )
            
            fig_profile = px.line(
                df_profile_melted,
                x='Hour of Day',
                y='Average Active Power (W)',
                color='Load Source',
                title="24-Hour Average Total Load Profile (W)",
                markers=True, 
                color_discrete_map={'Total Load (W)': 'red'} # Highlight total load
            )

            fig_profile.update_layout(
                xaxis_title="Hour of Day",
                yaxis_title="Average Active Power (W)",
                hovermode="x unified",
                legend_title="Load Source"
            )
            
            st.plotly_chart(fig_profile, use_container_width=True)
            
            st.subheader("Raw Average Data Table")
            st.dataframe(avg_load_profile_df, use_container_width=True)
            
        else:
            st.warning("Could not calculate the Average Load Profile. Ensure raw power data was valid.")
else:
    st.info("⬆️ Please upload your three MSB data files in the sidebar to begin the analysis. The application is now running and awaiting files.")
