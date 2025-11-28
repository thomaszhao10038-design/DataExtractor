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
""")

# --- Constants for Data Processing ---
PSUM_OUTPUT_NAME = 'PSum (W)' 

# Mapping user-friendly format to Python's datetime format strings
DATE_FORMAT_MAP = {
    "DD/MM/YYYY": "%d/%m/%Y %H:%M:%S",
    "YYYY-MM-DD": "%Y-%m-%d %H:%M:%S"
}

# --- Function to Process Data ---
def process_uploaded_files(uploaded_files, file_configs):
    """
    Reads multiple CSV files, extracts configured columns, cleans PSum data, 
    and returns a dictionary of DataFrames based on individual file configurations.
    """
    processed_data = {}
    
    for i, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name
        config = file_configs[i] # Get the specific configuration for this file
        
        try:
            # Convert user-defined column letters to 0-based indices
            date_col_index = excel_col_to_index(config['date_col_str'])
            time_col_index = excel_col_to_index(config['time_col_str'])
            ps_um_col_index = excel_col_to_index(config['psum_col_str'])
            
            # Define the columns to extract for this file
            columns_to_extract = {
                date_col_index: 'Date',
                time_col_index: 'Time',
                ps_um_col_index: PSUM_OUTPUT_NAME
            }
            col_indices = list(columns_to_extract.keys())
            
            # Check for unique indices
            if len(set(col_indices)) != 3:
                st.error(f"Error for file **{filename}**: Date, Time, and PSum must be extracted from three unique column indices. Check columns {config['date_col_str']}, {config['time_col_str']}, {config['psum_col_str']}.")
                continue
                
            header_index = int(config['start_row_num']) - 1 # 0-based index for Pandas header argument
            date_format_string = DATE_FORMAT_MAP.get(config['selected_date_format'])
            separator = config['delimiter_input']
            
            # 1. Read the CSV using the specified settings
            df_full = pd.read_csv(
                uploaded_file, 
                header=header_index, 
                encoding='ISO-8859-1', 
                low_memory=False,
                sep=separator # Use the file's selected separator
            )
            
            # 2. Check if DataFrame has enough columns
            max_index = max(col_indices)
            if df_full.shape[1] < max_index + 1:
                 st.error(f"File **{filename}** failed to read data correctly. It only has {df_full.shape[1]} columns. This usually means the **CSV Delimiter** ('{separator}') is incorrect for this file.")
                 continue

            # 3. Extract only the required columns by their index (iloc)
            df_extracted = df_full.iloc[:, col_indices].copy()
            
            # 4. Rename the columns to the final names for output
            temp_cols = {
                k: v for k, v in columns_to_extract.items()
            }
            df_extracted.columns = temp_cols.values()
            
            # 5. Data Cleaning: Convert PSum to numeric, handling potential errors
            if PSUM_OUTPUT_NAME in df_extracted.columns:
                df_extracted[PSUM_OUTPUT_NAME] = pd.to_numeric(
                    df_extracted[PSUM_OUTPUT_NAME], 
                    errors='coerce' # Convert non-numeric values to NaN
                )

            # 6. Format Date and Time columns separately after parsing for correction
            combined_dt_str = df_extracted['Date'].astype(str) + ' ' + df_extracted['Time'].astype(str)

            datetime_series = pd.to_datetime(
                combined_dt_str, 
                errors='coerce',
                format=date_format_string 
            )
            
            # --- CHECK: Verify successful datetime parsing ---
            valid_dates_count = datetime_series.count()
            if valid_dates_count == 0:
                st.warning(f"File **{filename}**: No valid dates could be parsed. Check the 'Date Format for Parsing' setting (**{config['selected_date_format']}**) and ensure the 'Date' and 'Time' columns contain valid data starting from Row {config['start_row_num']}.")
                continue
            # ---------------------------------------------------

            # GUARANTEE SEPARATION: Create a new DataFrame explicitly with separated columns
            df_final = pd.DataFrame({
                'Date': datetime_series.dt.strftime('%d/%m/%Y'), # Output Date is consistently DD/MM/YYYY
                'Time': datetime_series.dt.strftime('%H:%M:%S'),
                PSUM_OUTPUT_NAME: df_extracted[PSUM_OUTPUT_NAME] # Keep the PSum data from the original extracted DF
            })

            # 7. Clean the filename for the Excel sheet name
            sheet_name = filename.replace('.csv', '').replace('.', '_').strip()[:31]
            
            # Use the new, explicitly constructed DataFrame for the output
            processed_data[sheet_name] = df_final
            
        except ValueError as e:
            st.error(f"Configuration Error for file **{filename}**: Invalid column letter entered: {e}. Please use valid Excel column notation (e.g., A, C, AA).")
            continue
        except Exception as e:
            # Catch all other unexpected exceptions
            st.error(f"Error processing file **{filename}**. An unexpected error occurred. Error: {e}")
            continue
            
    return processed_data


# --- Function to Generate Excel File for Download ---
@st.cache_data
def to_excel(data_dict):
    """
    Takes a dictionary of DataFrames and writes them to an in-memory Excel file.
    The index is NOT included (index=False).
    
    Explicitly sets column formats to text using xlsxwriter to prevent merging 
    of Date and Time columns by Excel.
    """
    output = BytesIO()
    # Use pandas ExcelWriter with 'xlsxwriter' engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # --- Explicitly set column formats to Text (Crucial Fix) ---
            
            # Get the xlsxwriter workbook and worksheet objects.
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Define a text format (num_format: '@')
            text_format = workbook.add_format({'num_format': '@'})
            
            # Find column indices and apply the text format
            try:
                if 'Date' in df.columns:
                    date_col_index = df.columns.get_loc('Date')
                    # Apply text format to the entire column
                    worksheet.set_column(date_col_index, date_col_index, 12, text_format) 
                
                if 'Time' in df.columns:
                    time_col_index = df.columns.get_loc('Time')
                    # Apply text format to the entire column
                    worksheet.set_column(time_col_index, time_col_index, 10, text_format)
            except Exception as e:
                # Log any errors during explicit formatting but don't stop execution
                print(f"Error applying explicit xlsxwriter formats: {e}")
            # --------------------------------------------------------
            
    output.seek(0)
    return output.getvalue()


# --- Main Streamlit Logic ---
if __name__ == "__main__":
    
    # File Uploader is in the main area now
    uploaded_files = st.file_uploader(
        "Choose up to 10 CSV files", 
        type=["csv"], 
        accept_multiple_files=True
    )
    
    # Limit to 10 files
    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"You have uploaded {len(uploaded_files)} files. Only the first 10 will be processed.")
        uploaded_files = uploaded_files[:10]

    # Dynamic Configuration Section (appears only after files are uploaded)
    if uploaded_files:
        st.header("Individual File Configuration")
        st.warning("Please verify the Delimiter, Start Row, and Column Letters for each file below. Files with different delimiters must be configured separately.")
        
        file_configs = []
        all_configs_valid = True

        for i, uploaded_file in enumerate(uploaded_files):
            # Use a Streamlit expander for a cleaner interface for multiple files
            with st.expander(f"⚙️ Settings for **{uploaded_file.name}**", expanded=i == 0):
                
                # --- COLUMN CONFIGURATION (for this file) ---
                st.subheader("Column Letters")
                date_col_str = st.text_input(
                    "Date Column Letter", 
                    value='A', 
                    key=f'date_col_str_{i}'
                )

                time_col_str = st.text_input(
                    "Time Column Letter", 
                    value='B', 
                    key=f'time_col_str_{i}'
                )

                ps_um_col_str = st.text_input(
                    "PSum Column Letter", 
                    value='BI', 
                    key=f'psum_col_str_{i}',
                    help="PSum (Total Active Power) column letter in this file (e.g., 'BI')."
                )

                # --- CSV FILE SETTINGS (for this file) ---
                st.subheader("CSV File Parsing")
                delimiter_input = st.text_input(
                    "CSV Delimiter (Separator)",
                    value=',',
                    key=f'delimiter_input_{i}',
                    help="The character used to separate values (e.g., ',', ';', or '\\t')."
                )

                start_row_num = st.number_input(
                    "Header Row Number",
                    min_value=1,
                    value=3, 
                    key=f'start_row_num_{i}',
                    help="The row number that contains the column headers (e.g., 'Date', 'Time', 'UA')." 
                )

                selected_date_format = st.selectbox(
                    "Date Format for Parsing",
                    options=["DD/MM/YYYY", "YYYY-MM-DD"],
                    index=0, 
                    key=f'selected_date_format_{i}',
                    help="The date format in this file's Date column."
                )
                
                # Store the configuration dictionary for this file
                config = {
                    'date_col_str': date_col_str,
                    'time_col_str': time_col_str,
                    'psum_col_str': ps_um_col_str,
                    'delimiter_input': delimiter_input,
                    'start_row_num': start_row_num,
                    'selected_date_format': selected_date_format,
                }
                file_configs.append(config)
                
        # --- Processing and Download Button ---
        if st.button("🚀 Process All Files"):
            
            # 1. Process data 
            processed_data_dict = process_uploaded_files(
                uploaded_files, 
                file_configs
            )
            
            if processed_data_dict:
                
                # --- CONSOLIDATED RAW DATA SECTION ---
                st.header("Consolidated Raw Data Output")
                
                # Display a preview of the first processed file
                first_sheet_name = next(iter(processed_data_dict))
                st.subheader(f"Preview of: {first_sheet_name}")
                st.dataframe(processed_data_dict[first_sheet_name].head())
                st.success(f"Successfully processed {len(processed_data_dict)} of {len(uploaded_files)} file(s).")

                # --- File Name Customization ---
                file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
                
                if len(file_names_without_ext) > 1:
                    first_name = file_names_without_ext[0]
                    if len(first_name) > 20:
                        first_name = first_name[:17] + "..."
                    default_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
                elif file_names_without_ext:
                    default_filename = f"{file_names_without_ext[0]}_Consolidated.xlsx"
                else:
                    default_filename = "EnergyAnalyser_Consolidated_Data.xlsx"


                custom_filename = st.text_input(
                    "Output Excel Filename:",
                    value=default_filename,
                    key="output_filename_input_raw",
                    help="Enter the name for the final Excel file with raw extracted data."
                )
                
                # Generate Excel file for raw data
                excel_data = to_excel(processed_data_dict)
                
                # Download Button for raw data
                st.download_button(
                    label="📥 Download Consolidated Data (Date, Time, PSum)",
                    data=excel_data,
                    file_name=custom_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Click to download the Excel file with one sheet per uploaded CSV file."
                )

            else:
                st.error("No data could be successfully processed. Please review the error messages above and adjust the configurations in the file settings.")
    else:
        st.sidebar.markdown("Upload files to configure settings.")
