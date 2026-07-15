import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="UREA Shift Data Sync", layout="centered")

st.title("Daily UREA Plant Data Sync")
st.markdown("Upload your daily analytical report to auto-populate the DAR template.")

# Load the mapping file and cache it so it doesn't reload on every interaction
@st.cache_data
def load_mapping():
    return pd.read_excel("Book1.xlsx")

try:
    mapping_df = load_mapping()
except Exception as e:
    st.error(f"Error loading mapping file (Book1.xlsx). Make sure it is in the repository. Error: {e}")
    st.stop()

# Uploader for the daily source file
source_file = st.file_uploader("Upload Daily Log (must contain 'UREA' sheet)", type=["xlsx", "xls"])

if source_file is not None:
    try:
        # Load the uploaded source file
        source_wb = openpyxl.load_workbook(source_file, data_only=True)
        
        if "UREA" in source_wb.sheetnames:
            source_sheet = source_wb["UREA"]
            
            # Extract date from A1
            file_date = source_sheet['A1'].value
            st.success(f"✅ Daily log loaded successfully. Detected Date: **{file_date}**")
            
            # Load the target DAR master template from the repository
            target_filename = "DAR 27-06-2026.xlsx"
            target_wb = openpyxl.load_workbook(target_filename)
            target_sheet = target_wb["Sheet1"] 
            
            # ---------------------------------------------------------
            # DATA MAPPING LOGIC
            # Loop through Book1.xlsx and transfer values
            # ---------------------------------------------------------
            for index, row in mapping_df.iterrows():
                source_cell = row['Source_Cell']
                target_cell = row['Target_Cell']
                
                # Check if the cells are not empty in the mapping file
                if pd.notna(source_cell) and pd.notna(target_cell):
                    target_sheet[target_cell].value = source_sheet[source_cell].value
            
            # Save the updated target file into an in-memory buffer
            output = BytesIO()
            target_wb.save(output)
            output.seek(0)
            
            st.info("Template populated successfully! Ready for download.")
            
            # Format the date nicely for the filename if it exists
            if file_date:
                safe_date = str(file_date).replace("/", "-").replace(":", "-").split()[0]
            else:
                safe_date = "Updated"
                
            # Provide the download button for the new file
            st.download_button(
                label="📥 Download Updated DAR Report",
                data=output,
                file_name=f"DAR_{safe_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        else:
            st.error("The uploaded file does not contain a sheet named 'UREA'. Please verify the file.")
            
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
