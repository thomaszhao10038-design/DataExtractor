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
st.title("⚡ EnergyAnalyser: Data Consolidation")
st.markdown("""
    Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** and consolidate them into a single Excel file.
    
    The application uses the third row (index 2) of the CSV file as the main column header.
""")

# --- Constants for Data Processing ---
HEADER_ROW_INDEX = 2 # The 3rd row (Date, Time, UA, UB, etc.)

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
                 # Inform the user what the 1-based index (column letter) was that failed
                 st.error(f"File **{filename}** has only {df_full.shape[1]} columns. The column requested ({columns_config[max_index]} at index {max_index + 1}) is out of bounds. Please check the file structure or adjust the column letter in the sidebar.")
                 continue

            # 3. Extract only the required columns by their index (iloc)
            df_extracted = df_full.iloc[:, col_indices].copy()
            
            # 4. Rename the columns to the user-specified names
            df_extracted.columns = columns_config.values()
            
            # 5. Data Cleaning: Convert PSum to numeric, handling potential errors
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
    
    # Try to convert column letters to 0-based indices
    try:
        date_col_index = excel_col_to_index(date_col_str)
        time_col_index = excel_col_to_index(time_col_str)
        ps_um_col_index = excel_col_to_index(ps_um_col_str)
        
        COLUMNS_TO_EXTRACT = {
            date_col_index: 'Date',
            time_col_index: 'Time',
            ps_um_col_index: 'PSum'
        }
        
    except ValueError as e:
        st.error(f"Configuration Error: Invalid column letter entered: {e}. Please use valid Excel column notation (e.g., A, C, AA).")
        # Use st.stop() to prevent the rest of the app from running with invalid config
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
        
        # 1. Process data
        processed_data_dict = process_uploaded_files(uploaded_files, COLUMNS_TO_EXTRACT, HEADER_ROW_INDEX)
        
        if processed_data_dict:
            
            # Display a preview of the first processed file
            first_sheet_name = next(iter(processed_data_dict))
            st.subheader(f"Preview of: {first_sheet_name}")
            st.dataframe(processed_data_dict[first_sheet_name].head())
            st.success("All selected columns extracted and consolidated successfully!")
            
            # --- File Name Customization (Updated Logic) ---
            # Generate a dynamic default filename based on uploaded files
            file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
            
            if len(file_names_without_ext) > 1:
                # Use a combined name for multiple files
                first_name = file_names_without_ext[0]
                # Keep the default filename manageable
                if len(first_name) > 20:
                    first_name = first_name[:17] + "..."
                    
                default_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
            elif file_names_without_ext:
                # Use the single file name
                default_filename = f"{file_names_without_ext[0]}_Consolidated.xlsx"
            else:
                # Fallback
                default_filename = "EnergyAnalyser_Consolidated_Data.xlsx"


            custom_filename = st.text_input(
                "Output Excel Filename:", # Label updated
                value=default_filename,
                key="output_filename_input",
                help="Enter the name for the final Excel file, ensuring it ends with .xlsx"
            )
            
            # 2. Generate Excel file
            excel_data = to_excel(processed_data_dict)
            
            # 3. Download Button
            st.download_button(
                label="📥 Download Consolidated Data as Excel",
                data=excel_data,
                file_name=custom_filename, # Use the customized filename
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the Excel file with one sheet per uploaded CSV file."
            )
            
        else:
            st.error("No data could be successfully processed. Please review the error messages above and adjust the column letters if necessary.")
