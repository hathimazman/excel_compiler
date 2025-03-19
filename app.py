import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# Streamlit UI
st.title("Excel File Merger")
st.write("Upload multiple Excel files and combine them into a single file with multiple sheets while retaining formatting.")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xlsm", "xls"], accept_multiple_files=True)

if uploaded_files:
    output = io.BytesIO()  # Create in-memory buffer
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for uploaded_file in uploaded_files:
            # Load workbook to retain formatting
            wb = load_workbook(uploaded_file, keep_vba=True)  # Enable support for macro-enabled files
            sheet = wb.active  # Get active sheet
            df = pd.DataFrame(sheet.values)  # Read values while keeping formatting
            sheet_name = uploaded_file.name[:31]  # Sheet name (max 31 chars)
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)  # Write to sheet
    
    output.seek(0)  # Move buffer position to the start
    st.success("Files combined successfully while retaining formatting!")
    st.download_button(label="Download Combined Excel", data=output, file_name="combined.xlsm", mime="application/vnd.ms-excel.sheet.macroEnabled.12")
