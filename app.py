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
    """
    col_str = col_str.upper().strip()
    index = 0
    for char in col_str:
        if 'A' <= char <= 'Z':
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    return index - 1  # Convert 1-based to 0-based

# --- App Title and Description ---
st.title("⚡ EnergyAnalyser: Data Consolidation")
st.markdown("""
Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** and consolidate them into a single Excel file.

The application uses the third row (index 2) of the CSV file as the main column header.
""")

# --- Constants ---
HEADER_ROW_INDEX = 2
PSUM_OUTPUT_NAME = 'PSum (W)'

# --- Sidebar Column Configuration ---
st.sidebar.header("⚙️ Column Configuration")
st.sidebar.markdown("Define the column letter for each data field.")

date_col_str = st.sidebar.text_input("Date Column Letter (Default: A)", value='A')
time_col_str = st.sidebar.text_input("Time Column Letter (Default: B)", value='B')
ps_um_col_str = st.sidebar.text_input("PSum Column Letter (Default: BI)", value='BI',
                                     help="PSum (Total Active Power) is expected in Excel column BI. Adjust if needed.")

# --- Function to Process Uploaded CSV Files ---
def process_uploaded_files(uploaded_files, columns_config, header_index):
    """
    Reads multiple CSV files, extracts configured columns, cleans PSum data,
    and returns a dictionary of DataFrames with separate Date and Time columns.
    """
    processed_data = {}

    # Ensure unique columns
    if len(set(columns_config.keys())) != 3:
        st.error("Date, Time, and PSum must be from three unique columns.")
        return {}

    col_indices = list(columns_config.keys())

    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        try:
            # Read CSV
            df_full = pd.read_csv(uploaded_file, header=header_index, encoding='ISO-8859-1', low_memory=False)

            # Check column existence
            max_index = max(col_indices)
            if df_full.shape[1] < max_index + 1:
                col_name = columns_config.get(max_index, 'Unknown')
                st.error(f"File **{filename}** has only {df_full.shape[1]} columns. Column requested ({col_name} at index {max_index + 1}) is out of bounds.")
                continue

            # Extract required columns
            df_extracted = df_full.iloc[:, col_indices].copy()
            df_extracted.columns = list(columns_config.values())

            # Convert PSum to numeric
            df_extracted[PSUM_OUTPUT_NAME] = pd.to_numeric(df_extracted[PSUM_OUTPUT_NAME], errors='coerce')

            # Separate Date and Time columns properly
            df_extracted['Date'] = pd.to_datetime(df_extracted['Date'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
            df_extracted['Time'] = pd.to_datetime(df_extracted['Time'], errors='coerce').dt.strftime('%H:%M:%S')

            # Keep only needed columns
            df_final = df_extracted[['Date', 'Time', PSUM_OUTPUT_NAME]].copy()

            # Clean filename for Excel sheet
            sheet_name = filename.replace('.csv', '').replace('.', '_').strip()[:31]

            processed_data[sheet_name] = df_final

        except Exception as e:
            st.error(f"Error processing file **{filename}**. Error: {e}")
            continue

    return processed_data

# --- Function to Convert DataFrames to Excel ---
@st.cache_data
def to_excel(data_dict):
    """
    Converts a dictionary of DataFrames to an in-memory Excel file.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# --- Main Streamlit Logic ---
if __name__ == "__main__":

    # Convert column letters to 0-based indices
    try:
        date_col_index = excel_col_to_index(date_col_str)
        time_col_index = excel_col_to_index(time_col_str)
        ps_um_col_index = excel_col_to_index(ps_um_col_str)

        COLUMNS_TO_EXTRACT = {
            date_col_index: 'Date',
            time_col_index: 'Time',
            ps_um_col_index: PSUM_OUTPUT_NAME
        }

    except ValueError as e:
        st.error(f"Configuration Error: {e}. Use valid Excel letters (A, C, AA).")
        st.stop()

    # Upload CSV files
    uploaded_files = st.file_uploader(
        "Choose up to 10 CSV files", 
        type=["csv"], 
        accept_multiple_files=True
    )

    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"You uploaded {len(uploaded_files)} files. Only the first 10 will be processed.")
        uploaded_files = uploaded_files[:10]

    # Process files
    if uploaded_files:
        st.info(f"Processing {len(uploaded_files)} file(s) using columns: Date: {date_col_str.upper()}, Time: {time_col_str.upper()}, PSum: {ps_um_col_str.upper()}.")

        processed_data_dict = process_uploaded_files(uploaded_files, COLUMNS_TO_EXTRACT, HEADER_ROW_INDEX)

        if processed_data_dict:
            st.header("Consolidated Raw Data Output")

            first_sheet_name = next(iter(processed_data_dict))
            st.subheader(f"Preview of: {first_sheet_name}")
            st.dataframe(processed_data_dict[first_sheet_name].head())
            st.success("Selected columns extracted and consolidated successfully!")

            # Default filename
            file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
            if len(file_names_without_ext) > 1:
                first_name = file_names_without_ext[0][:17] + "..." if len(file_names_without_ext[0]) > 20 else file_names_without_ext[0]
                default_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
            else:
                default_filename = f"{file_names_without_ext[0]}_Consolidated.xlsx" if file_names_without_ext else "EnergyAnalyser_Consolidated_Data.xlsx"

            custom_filename = st.text_input("Output Excel Filename:", value=default_filename)

            # Generate Excel
            excel_data = to_excel(processed_data_dict)

            # Download Button
            st.download_button(
                label="📥 Download Consolidated Data (Date, Time, PSum)",
                data=excel_data,
                file_name=custom_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No data could be successfully processed. Please check the error messages above.")
