import streamlit as st
import pandas as pd
import numpy as np

# --- Configuration ---
# Based on the file snippets for 'final document.xlsx - MSB X.csv', the files seem to 
# have complex headers that repeat. We will try a simplified approach by targeting 
# the power column directly, assuming the file format is consistent.
# For the MSB X.csv files, the actual column names appear in the first row.
HEADER_ROW_INDEX = 0 
# Required columns based on the file snippets for MSB 1, 2, 3:
REQUIRED_TIME_COLUMN = 'Local Time Stamp'
REQUIRED_POWER_COLUMN = 'Active Power (W)'
REQUIRED_COLUMNS = [REQUIRED_TIME_COLUMN, REQUIRED_POWER_COLUMN]

def find_first_occurrence(df, column_name):
    """
    Finds the index of the first occurrence of a column name,
    which is necessary when the CSV has duplicate column headers.
    """
    cols = df.columns.tolist()
    try:
        # Find the index of the first occurrence of the column name
        index = cols.index(column_name)
        return df.columns[index]
    except ValueError:
        return None

def process_msb_data(file_path, msb_name):
    """
    Reads and processes the MSB raw data files.
    This function is corrected to use column names and the appropriate header index.
    """
    st.write(f"--- Processing {msb_name} Data ---")
    
    try:
        # Use header=HEADER_ROW_INDEX (0) to correctly set the column names 
        # based on the row that starts with 'UTC Offset (minutes), Local Time Stamp, Active Power (W)...'.
        # Since the columns are repeated, we will load the whole file and select by index.
        df_raw = pd.read_csv(file_path, header=HEADER_ROW_INDEX)
        
        # --- Column Selection (Handling Duplicate Columns) ---
        
        # Find the first 'Local Time Stamp' column (it's the second column in the row 0 header)
        time_col_name = df_raw.columns[1]
        
        # Find the first 'Active Power (W)' column (it's the third column in the row 0 header)
        power_col_name = df_raw.columns[2]

        if not (time_col_name == REQUIRED_TIME_COLUMN and power_col_name == REQUIRED_POWER_COLUMN):
            st.error(f"Error for {msb_name}: Expected columns '{REQUIRED_TIME_COLUMN}' (Col 1) and '{REQUIRED_POWER_COLUMN}' (Col 2) not found at expected positions.")
            st.warning(f"Columns found at [1] and [2]: {time_col_name}, {power_col_name}")
            return None
        
        # Select the necessary columns by index 1 and 2 to grab the first occurrence of Time and Power
        df = df_raw.iloc[:, [1, 2]].copy()
        
        # Apply the required column names for consistency
        df.columns = [REQUIRED_TIME_COLUMN, REQUIRED_POWER_COLUMN]
        
        # Drop the first row which contains the initial timestamp in the header data
        df = df.iloc[1:].copy()
        
        # Rename and process
        df.rename(columns={REQUIRED_POWER_COLUMN: 'Power (W)', REQUIRED_TIME_COLUMN: 'Timestamp'}, inplace=True)
        
        # Convert power from W to kW
        # Also convert column to numeric, as it may be read as string due to missing values
        df['Power (W)'] = pd.to_numeric(df['Power (W)'], errors='coerce')
        df['Power (kW)'] = df['Power (W)'] / 1000.0
        
        # Remove negative power values (likely export/measurement error for simple load analysis)
        df['Power (kW)'] = df['Power (kW)'].clip(lower=0) 
        
        st.success(f"Successfully loaded and processed {msb_name}. Rows: {len(df)}")
        st.dataframe(df.head())
        
        return df.dropna(subset=['Power (kW)', 'Timestamp'])

    except FileNotFoundError:
        st.error(f"File not found for {msb_name} at path: {file_path}")
        return None
    except Exception as e:
        # Catch any other unexpected error and report it clearly
        st.error(f"An unexpected error occurred while processing {msb_name}: {e}")
        st.info("This often happens when the file structure is highly complex or inconsistent.")
        return None

def main():
    """Main Streamlit application logic."""
    st.set_page_config(layout="wide", page_title="Energy Data Analyser (MSB)")
    st.title("Energy Data Analyser (MSB)")
    st.markdown("Load and analyze raw energy data from MSB devices.")

    # --- File Paths (Updated to the structure of the newer, raw-looking MSB files) ---
    msb_files = {
        'MSB-1': 'final document.xlsx - MSB 1.csv',
        'MSB-2': 'final document.xlsx - MSB 2.csv',
        'MSB-3': 'final document.xlsx - MSB 3.csv',
    }
    
    # --- Data Processing and Visualization ---
    
    all_dfs = {}
    
    for msb_name, file_path in msb_files.items():
        df = process_msb_data(file_path, msb_name)
        if df is not None:
            all_dfs[msb_name] = df
            
    if all_dfs:
        st.header("Combined Power Overview (kW)")
        st.markdown("This chart shows the total active power demand for each MSB over time.")

        combined_data = pd.DataFrame()
        
        # Plotting the data if processing was successful
        for msb_name, df in all_dfs.items():
            if 'Power (kW)' in df.columns and 'Timestamp' in df.columns:
                # Set Timestamp as index and add MSB name as column header
                df_plot = df.set_index('Timestamp')['Power (kW)'].rename(msb_name)
                combined_data = pd.concat([combined_data, df_plot], axis=1)

        if not combined_data.empty:
            # We are not dropping NaNs here to preserve data points where one MSB 
            # might have a value while others don't, but the line chart handles this.
            
            st.line_chart(combined_data)
            
            # --- Aggregated Metrics ---
            total_load = combined_data.sum(axis=1)
            total_load_summary = {
                'Total Peak Demand (kW)': total_load.max(),
                'Average Load (kW)': total_load.mean(),
                'Total Energy (kWh)': total_load.sum() / (60 / 5) # Assuming 5-minute intervals (12 data points per hour)
            }
            
            st.subheader("Aggregated Metrics")
            st.dataframe(pd.DataFrame(total_load_summary, index=['Value']))

            st.subheader("Data Summary (First 500 rows)")
            st.dataframe(combined_data.head(500))
        else:
            st.info("No common timestamp data available to combine and display.")
            
if __name__ == "__main__":
    main()
