import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuration for Streamlit Page ---
st.set_page_config(
    page_title="EnergyAnalyser",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- App Title and Description ---
st.title("⚡ EnergyAnalyser: Data Consolidation")
st.markdown("""
    Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** and consolidate them into a single Excel file.
    
    The application uses the third row (index 2) of the CSV file as the main column header.
""")

# --- Constants for Data Processing ---
HEADER_ROW_INDEX = 2 # The 3rd row (Date, Time, UA, UB, etc.)

# --- User Configuration Section (Sidebar) ---
st.sidebar.header("⚙️ Column Index Configuration")
st.sidebar.markdown("Define the column index for each required data field (0-based).")

# Get user-defined indices from the sidebar, with robust defaults
date_col_index = st.sidebar.number_input(
    "Date Column Index (Default: 0)", 
    min_value=0, 
    value=0, 
    step=1, 
    key='date_idx'
)

time_col_index = st.sidebar.number_input(
    "Time Column Index (Default: 1)", 
    min_value=0, 
    value=1, 
    step=1, 
    key='time_idx'
)

ps_um_col_index = st.sidebar.number_input(
    "PSum Column Index (Total Active Power) (Default: 40)", 
    min_value=0, 
    value=40, 
    step=1, 
    key='psum_idx',
    help="This is the 0-based column index. If the PSum data is in Excel column AO, the index is 40 (A=0, B=1, ... AO=40)."
)

# Define the columns to extract using the user-configured values
COLUMNS_TO_EXTRACT = {
    date_col_index: 'Date',
    time_col_index: 'Time',
    ps_um_col_index: 'PSum'
}

# --- Function to Process Data ---
def process_uploaded_files(uploaded_files, columns_config, header_index):
    """
    Reads multiple CSV files, extracts configured columns, cleans PSum data, 
    and returns a dictionary of DataFrames.
    """
    processed_data = {}
    
    # Ensure all required columns are unique
    if len(set(columns_config.keys())) != 3:
        st.error("Error: Date, Time, and PSum must be extracted from three unique column indices.")
        return {}

    col_indices = list(columns_config.keys())
    
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        
        try:
            # 1. Read the CSV using the specified header row
            # header=2: Use the 3rd row as column names
            df_full = pd.read_csv(
                uploaded_file, 
                header=header_index, 
                encoding='ISO-8859-1', 
                low_memory=False
            )
            
            # 2. Check if DataFrame has enough columns
            max_index = max(col_indices)
            if df_full.shape[1] < max_index + 1:
                 st.error(f"File **{filename}** has only {df_full.shape[1]} columns, but column index **{max_index}** was requested. Please check the file structure or adjust the indices in the sidebar.")
                 continue

            # 3. Extract only the required columns by their index (iloc)
            df_extracted = df_full.iloc[:, col_indices].copy()
            
            # 4. Rename the columns to the user-specified names
            df_extracted.columns = columns_config.values()
            
            # 5. Data Cleaning: Convert PSum to numeric, handling potential errors
            # We assume the user has configured the correct PSum column index
            if 'PSum' in df_extracted.columns:
                df_extracted['PSum'] = pd.to_numeric(
                    df_extracted['PSum'], 
                    errors='coerce' # Convert non-numeric values to NaN
                )
            
            # 6. Clean the filename for the Excel sheet name
            sheet_name = filename.replace('.csv', '').replace('.', '_').strip()[:31]
            
            processed_data[sheet_name] = df_extracted
            
        except Exception as e:
            st.error(f"Error processing file **{filename}**. An unexpected error occurred. Error: {e}")
            continue
            
    return processed_data

# --- Function to Generate Excel File for Download ---
@st.cache_data
def to_excel(data_dict):
    """
    Takes a dictionary of DataFrames and writes them to an in-memory Excel file.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    return output.getvalue()


# --- Main Streamlit Logic ---
if __name__ == "__main__":
    
    # File Uploader
    uploaded_files = st.file_uploader(
        "Choose up to 10 CSV files", 
        type=["csv"], 
        accept_multiple_files=True
    )
    
    # Limit to 10 files
    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"You have uploaded {len(uploaded_files)} files. Only the first 10 will be processed.")
        uploaded_files = uploaded_files[:10]

    # Processing and Download Button
    if uploaded_files:
        
        st.info(f"Processing {len(uploaded_files)} file(s). Check the sidebar to confirm column indices are set correctly (Date: {date_col_index}, Time: {time_col_index}, PSum: {ps_um_col_index}).")
        
        # 1. Process data
        processed_data_dict = process_uploaded_files(uploaded_files, COLUMNS_TO_EXTRACT, HEADER_ROW_INDEX)
        
        if processed_data_dict:
            
            # Display a preview of the first processed file
            first_sheet_name = next(iter(processed_data_dict))
            st.subheader(f"Preview of: {first_sheet_name}")
            st.dataframe(processed_data_dict[first_sheet_name].head())
            st.success("All selected columns extracted and consolidated successfully!")
            
            # 2. Generate Excel file
            excel_data = to_excel(processed_data_dict)
            
            # 3. Download Button
            st.download_button(
                label="📥 Download Consolidated Data as Excel",
                data=excel_data,
                file_name="EnergyAnalyser_Consolidated_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the Excel file with one sheet per uploaded CSV file."
            )
            
        else:
            st.error("No data could be successfully processed. Please review the error messages above and adjust the column indices if necessary.")
