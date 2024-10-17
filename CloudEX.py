import streamlit as st
import pandas as pd
import openpyxl
import shutil
import datetime
import os
import json
from openpyxl.utils import get_column_letter

# Get the current directory of the script
current_dir = os.path.dirname(os.path.abspath(__file__))

# File paths (use relative paths now)
TEMPLATE_PATH = os.path.join(current_dir, 'BC CALC.xlsx')
LOG_PATH = os.path.join(current_dir, 'generated_files_log.json')

# Streamlit app
st.set_page_config(page_title="Excel Data Transfer Bot", layout="wide")

# Title
st.title('Excel Data Transfer Bot')

# Step 1: User inputs keyword
keyword = st.text_input('Enter the keyword for the output filename:', key='keyword_input')

# Step 2: User uploads CSV file
uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"], key='file_upload')

# Step 3: Generate button for user to initiate generation
if st.button("Generate", key='generate_button'):
    if not keyword:
        st.warning("Please enter a keyword before generating.")
    elif not uploaded_file:
        st.warning("Please upload a CSV file before generating.")
    else:
        try:
            # Extract today's date for the output file name
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            output_filename = f"{keyword}_{today}.xlsx"
            output_filepath = os.path.join(current_dir, output_filename)

            # Step 4: Load the input CSV file into a pandas DataFrame
            input_df = pd.read_csv(uploaded_file)
            
            # Step 5: Copy the template file to create a new file
            shutil.copy(TEMPLATE_PATH, output_filepath)
            
            # Step 6: Open the new file using openpyxl
            workbook = openpyxl.load_workbook(output_filepath)
            worksheet = workbook.active
            
            # Helper function to extract data based on column name
            def extract_data(df, possible_column_names, start_row, column_letter):
                for column_name in possible_column_names:
                    matching_columns = [col for col in df.columns if column_name in col.lower()]
                    if matching_columns:
                        data = df[matching_columns[0]].iloc[:10].tolist()  # Extract first 10 rows under the header
                        for i, value in enumerate(data):
                            worksheet[f"{column_letter}{start_row + i}"] = value
                        return
            
            # Step 7: Extract and write 'Product Details'
            extract_data(input_df, ['product details', 'product'], 4, 'F')
            
            # Step 8: Extract and write 'Brand'
            extract_data(input_df, ['brand'], 4, 'G')
            
            # Step 9: Extract and write 'Price'
            extract_data(input_df, ['price'], 4, 'H')
            
            # Step 10: Extract and write 'Revenue'
            extract_data(input_df, ['revenue'], 4, 'I')
            
            # Step 11: Correct 'Revenue' cells to be numeric and format as currency
            for i in range(4, 14):
                cell = worksheet[f"I{i}"]
                if cell.value is not None:
                    try:
                        # Remove commas from the cell value to allow conversion to float
                        cleaned_value = str(cell.value).replace(",", "")
                        numeric_value = float(cleaned_value)
                        
                        # Set the cell's value to the numeric value
                        cell.value = numeric_value
                        
                        # Apply the desired number format (e.g., Currency)
                        cell.number_format = '$#,##0.00'  # You can customize this format as needed
                    except ValueError:
                        st.warning(f"Cell I{i} contains non-numeric data: {cell.value}")
            
            # Save the updated workbook
            workbook.save(output_filepath)
            
            # Step 12: Update log file
            log_entry = {
                "keyword": keyword,
                "filename": output_filename,
                "timestamp": today
            }
            if os.path.exists(LOG_PATH):
                with open(LOG_PATH, "r") as log_file:
                    log_data = json.load(log_file)
            else:
                log_data = []
            log_data.append(log_entry)
            with open(LOG_PATH, "w") as log_file:
                json.dump(log_data, log_file, indent=4)
            
            # Step 13: Notify user of generation success
            st.success(f"File generated successfully: {output_filename}")
            
            # Provide download button
            with open(output_filepath, "rb") as file:
                st.download_button(label="Download Excel File", data=file, file_name=output_filename, key='download_button')
        except Exception as e:
            st.error(f"An error occurred: {e}")

# Log section
st.sidebar.title("Generated Files Log")
if os.path.exists(LOG_PATH):
    with open(LOG_PATH, "r") as log_file:
        log_data = json.load(log_file)
    for entry in reversed(log_data):
        st.sidebar.write(f"Keyword: {entry['keyword']} | File: {entry['filename']} | Date: {entry['timestamp']}")
else:
    st.sidebar.write("No files have been generated yet.")
