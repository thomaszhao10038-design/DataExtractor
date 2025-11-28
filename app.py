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
    Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** (Total Active Power) 
    and consolidate them into a single Excel file.
    
    This app is configured to correctly read your specific log format:
    1. It skips the first two metadata/header rows.
    2. It uses the third row (index 2) as the main column header.
    3. It extracts 'Date', 'Time', and the 'PSum' (Total Active Power) column.
""")

# --- Constants for Data Processing ---
# Based on the file structure:
# skip initial 2 rows, use row 2 (index 2) as header
HEADER_ROW_INDEX = 2
# Column indices for the specific data we want (0-based)
COLUMNS_TO_EXTRACT = {
    0: 'Date',     # First column is Date
    1: 'Time',     # Second column is Time
    39: 'PSum'     # The 40th column is Total Active Power(W), which you call PSum
}

# --- Function to Process Data ---
def process_uploaded_files(uploaded_files):
    """
    Reads multiple CSV files, extracts Date, Time, and PSum based on fixed indices,
    and returns a dictionary of DataFrames.
    """
    processed_data = {}
    
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        
        try:
            # 1. Read the CSV using the specified header row
            # header=2 means use the 3rd row (0-indexed) as the column names
            # skiprows=1 means skip the very first row (ProductSN...)
            # NOTE: For the raw file you uploaded, using header=2 handles the first two lines correctly.
            df_full = pd.read_csv(
                uploaded_file, 
                header=HEADER_ROW_INDEX, 
                skiprows=[0], # Skip the first row (ProductSN)
                encoding='ISO-8859-1' # Use a robust encoding to handle file variations
            )
            
            # 2. Extract only the required columns by their index
            # This is robust because we use fixed indices (0, 1, 39)
            
            # Create a dictionary for remapping columns
            col_map = {idx: name for idx, name in COLUMNS_TO_EXTRACT.items()}
            
            # Select the columns by their location (iloc)
            # The indices are applied after the header row has been read and applied
            df_extracted = df_full.iloc[:, list(col_map.keys())]
            
            # 3. Rename the columns to 'Date', 'Time', 'PSum'
            df_extracted.columns = col_map.values()
            
            # 4. Clean the filename for the Excel sheet name
            sheet_name = filename.split('.')[0][:31] # Max 31 chars for Excel sheet name
            
            processed_data[sheet_name] = df_extracted
            
        except Exception as e:
            st.error(f"Error processing file {filename}: {e}")
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
        
        st.info(f"Processing {len(uploaded_files)} file(s)...")
        
        # 1. Process data
        processed_data_dict = process_uploaded_files(uploaded_files)
        
        if processed_data_dict:
            
            # Display a preview of the first processed file
            first_sheet_name = next(iter(processed_data_dict))
            st.subheader(f"Preview of: {first_sheet_name}")
            st.dataframe(processed_data_dict[first_sheet_name].head())
            st.success("All selected columns extracted successfully!")
            
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
            st.error("No data could be processed. Please check the format of your uploaded CSV files.")
