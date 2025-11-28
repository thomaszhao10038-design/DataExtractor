import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuration for Streamlit Page ---
st.set_page_config(
    page_title="EnergyAnalyser",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Helper Function for Excel Column Conversion ---
def excel_col_to_index(col_str):
    """
    Converts an Excel column string (e.g., 'A', 'AA', 'BI') to a 0-based column index.
    Raises a ValueError if the string is invalid.
    """
    col_str = col_str.upper().strip()
    index = 0
    # A=1, B=2, ..., Z=26
    for char in col_str:
        if 'A' <= char <= 'Z':
            # Calculate the 1-based index (e.g., 'B' is 2, 'AA' is 27)
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    
    # Convert 1-based index to 0-based index for Pandas (A=0, B=1)
    return index - 1

# --- App Title and Description ---
st.title("⚡ EnergyAnalyser: Data Consolidation & Resampling")
st.markdown("""
    Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** and consolidate them into a single Excel file.
    
    The application uses the third row (index 2) of the CSV file as the main column header.
""")

# --- Constants for Data Processing ---
HEADER_ROW_INDEX = 2 # The 3rd row (Date, Time, UA, UB, etc.)
PSUM_OUTPUT_NAME = 'PSum (W)' 
RESAMPLED_PSUM_NAME = '|PSum| Avg (W)'

# --- User Configuration Section (Sidebar) ---
st.sidebar.header("⚙️ Column Configuration")
st.sidebar.markdown("Define the column letter for each data field.")

# Get user-defined column letters (A, B, C...)
date_col_str = st.sidebar.text_input(
    "Date Column Letter (Default: A)", 
    value='A', 
    key='date_col_str'
)

time_col_str = st.sidebar.text_input(
    "Time Column Letter (Default: B)", 
    value='B', 
    key='time_col_str'
)

ps_um_col_str = st.sidebar.text_input(
    "PSum Column Letter (Default: BI)", 
    value='BI', 
    key='psum_col_str',
    help="PSum (Total Active Power) is expected in Excel column BI. Adjust if needed."
)

# --- Function to Process Data ---
def process_uploaded_files(uploaded_files, columns_config, header_index):
    """
    Reads multiple CSV files, extracts configured columns, cleans PSum data, 
    and returns a dictionary of DataFrames indexed by Datetime.
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
            df_full = pd.read_csv(
                uploaded_file, 
                header=header_index, 
                encoding='ISO-8859-1', 
                low_memory=False
            )
            
            # 2. Check if DataFrame has enough columns
            max_index = max(col_indices)
            if df_full.shape[1] < max_index + 1:
                 col_name = columns_config.get(max_index, 'Unknown')
                 st.error(f"File **{filename}** has only {df_full.shape[1]} columns. The column requested ({col_name} at index {max_index + 1}) is out of bounds. Please check the file structure or adjust the column letter in the sidebar.")
                 continue

            # 3. Extract only the required columns by their index (iloc)
            df_extracted = df_full.iloc[:, col_indices].copy()
            
            # 4. Rename the columns temporarily to enable combine/clean
            temp_cols = {
                columns_config[k]: v for k, v in COLUMNS_TO_EXTRACT.items()
            }
            df_extracted.columns = temp_cols.values()
            
            # 5. Combine Date and Time into a single Datetime column
            df_extracted['Datetime'] = pd.to_datetime(
                df_extracted['Date'] + ' ' + df_extracted['Time'], 
                errors='coerce',
                format='mixed' # Use mixed format to handle various date/time styles
            )
            
            # 6. Set Datetime as index and clean up
            df_extracted = df_extracted.set_index('Datetime').dropna(subset=['Datetime'])
            df_extracted = df_extracted.drop(columns=['Date', 'Time'])
            
            # 7. Data Cleaning: Convert PSum to numeric, handling potential errors
            if PSUM_OUTPUT_NAME in df_extracted.columns:
                df_extracted[PSUM_OUTPUT_NAME] = pd.to_numeric(
                    df_extracted[PSUM_OUTPUT_NAME], 
                    errors='coerce' # Convert non-numeric values to NaN
                )
            
            # 8. Clean the filename for the Excel sheet name
            sheet_name = filename.replace('.csv', '').replace('.', '_').strip()[:31]
            
            processed_data[sheet_name] = df_extracted
            
        except Exception as e:
            st.error(f"Error processing file **{filename}**. An unexpected error occurred. Error: {e}")
            continue
            
    return processed_data

# --- New Function for Resampling ---
def resample_10min_modulus(processed_data_dict):
    """
    Takes a dictionary of DataFrames, calculates the modulus of PSum (W),
    and resamples the data to a 10-minute average.
    """
    resampled_data = {}
    
    for sheet_name, df in processed_data_dict.items():
        # Ensure the PSum column is numeric and exists
        if PSUM_OUTPUT_NAME in df.columns:
            # 1. Calculate modulus (absolute value)
            df[RESAMPLED_PSUM_NAME] = df[PSUM_OUTPUT_NAME].abs()
            
            # 2. Resample to 10-minute mean
            # '10T' specifies 10 minutes.
            df_resampled = df[RESAMPLED_PSUM_NAME].resample('10T').mean().to_frame()
            
            # 3. Drop rows where resampling resulted in NaN (e.g., end of data)
            df_resampled = df_resampled.dropna()
            
            resampled_data[sheet_name] = df_resampled
        # Note: We skip files where PSum wasn't found or wasn't numeric after cleaning
            
    return resampled_data


# --- Function to Generate Excel File for Download ---
@st.cache_data
def to_excel(data_dict):
    """
    Takes a dictionary of DataFrames and writes them to an in-memory Excel file.
    The index (Datetime) is included as the first column.
    """
    output = BytesIO()
    # Use pandas ExcelWriter to write to BytesIO
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            # index=True is necessary here because the Datetime is the index
            df.to_excel(writer, sheet_name=sheet_name, index=True)
    
    return output.getvalue()


# --- Main Streamlit Logic ---
if __name__ == "__main__":
    
    # Try to convert column letters to 0-based indices
    try:
        date_col_index = excel_col_to_index(date_col_str)
        time_col_index = excel_col_to_index(time_col_str)
        ps_um_col_index = excel_col_to_index(ps_um_col_str)
        
        # Define the columns to extract, using the constants
        COLUMNS_TO_EXTRACT = {
            date_col_index: 'Date',
            time_col_index: 'Time',
            ps_um_col_index: PSUM_OUTPUT_NAME
        }
        
    except ValueError as e:
        st.error(f"Configuration Error: Invalid column letter entered: {e}. Please use valid Excel column notation (e.g., A, C, AA).")
        st.stop()
        

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
        
        # Display the column letters being used for user confirmation
        st.info(f"Processing {len(uploaded_files)} file(s) using columns: Date: {date_col_str.upper()}, Time: {time_col_str.upper()}, PSum: {ps_um_col_str.upper()}.")
        
        # 1. Process data (Outputs time-indexed dataframes)
        processed_data_dict = process_uploaded_files(uploaded_files, COLUMNS_TO_EXTRACT, HEADER_ROW_INDEX)
        
        if processed_data_dict:
            
            # --- CONSOLIDATED RAW DATA SECTION ---
            st.header("1. Consolidated Raw Data Output")
            
            # Display a preview of the first processed file
            first_sheet_name = next(iter(processed_data_dict))
            st.subheader(f"Preview of: {first_sheet_name}")
            st.dataframe(processed_data_dict[first_sheet_name].head())
            st.success("Selected columns extracted and consolidated successfully!")
            
            # File Name Customization for raw data
            file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
            
            if len(file_names_without_ext) > 1:
                first_name = file_names_without_ext[0]
                if len(first_name) > 20:
                    first_name = first_name[:17] + "..."
                default_filename_raw = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
            elif file_names_without_ext:
                default_filename_raw = f"{file_names_without_ext[0]}_Consolidated.xlsx"
            else:
                default_filename_raw = "EnergyAnalyser_Consolidated_Data.xlsx"


            custom_filename_raw = st.text_input(
                "Output Excel Filename:",
                value=default_filename_raw,
                key="output_filename_input_raw",
                help="Enter the name for the final Excel file with raw extracted data."
            )
            
            # Generate Excel file for raw data
            excel_data_raw = to_excel(processed_data_dict)
            
            # Download Button for raw data
            st.download_button(
                label="📥 Download Consolidated Raw Data (Date, Time, PSum)",
                data=excel_data_raw,
                file_name=custom_filename_raw,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the Excel file with one sheet per uploaded CSV file."
            )

            # --- RESAMPLED DATA SECTION (NEW FEATURE) ---
            st.header("2. 10-Minute Resampled Modulus Output")
            
            # 1. Process for resampling
            resampled_data_dict = resample_10min_modulus(processed_data_dict)
            
            if resampled_data_dict:
                # Display a preview of the first resampled file
                first_sheet_name_resampled = next(iter(resampled_data_dict))
                st.subheader(f"Preview of: {first_sheet_name_resampled}")
                st.dataframe(resampled_data_dict[first_sheet_name_resampled].head())
                st.success("Data successfully resampled to 10-minute absolute mean!")
                
                # File Name Customization for resampled data
                # Use the raw filename base but adjust the suffix
                default_filename_resampled = default_filename_raw.replace("_Consolidated.xlsx", "_10min_Modulus.xlsx")
                default_filename_resampled = default_filename_resampled.replace("_More_Consolidated.xlsx", "_More_10min_Modulus.xlsx")


                custom_filename_resampled = st.text_input(
                    "Output Excel Filename (10-min Modulus):",
                    value=default_filename_resampled,
                    key="output_filename_input_resampled",
                    help="Enter the name for the final Excel file with 10-minute resampled data."
                )
                
                # Generate Excel file for resampled data
                excel_data_resampled = to_excel(resampled_data_dict)
                
                # Download Button for resampled data
                st.download_button(
                    label="📥 Download 10-Minute Modulus Data",
                    data=excel_data_resampled,
                    file_name=custom_filename_resampled,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Click to download the Excel file with the modulus of PSum averaged over 10-minute intervals."
                )
                
            else:
                st.warning("Could not resample data. Check if PSum (W) column contains valid numeric data after initial processing.")

        else:
            st.error("No data could be successfully processed. Please review the error messages above and adjust the column letters if necessary.")
