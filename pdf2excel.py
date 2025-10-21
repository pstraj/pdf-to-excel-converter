import streamlit as st
import pandas as pd
import tabula
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import subprocess
import sys

# Install java if not present (for Streamlit Cloud)
try:
    subprocess.run(['java', '-version'], capture_output=True, check=True)
except:
    st.error("Java is required but not found. Please contact the administrator.")

st.set_page_config(page_title="PDF to Excel Converter", layout="wide")

st.title("📄 PDF to Excel Converter")
st.write("Upload a PDF file, select and reorder columns, then download as Excel")

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'selected_columns' not in st.session_state:
    st.session_state.selected_columns = []
if 'column_order' not in st.session_state:
    st.session_state.column_order = []

# Step 1: File Upload
st.header("Step 1: Upload PDF File")
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    try:
        # Read PDF and extract tables
        with st.spinner("Reading PDF file..."):
            # Save uploaded file temporarily
            pdf_bytes = uploaded_file.read()
            
            # Use tabula to read PDF tables
            tables = tabula.read_pdf(io.BytesIO(pdf_bytes), pages='all', multiple_tables=True)
            
            if len(tables) == 0:
                st.error("No tables found in the PDF file.")
            else:
                # If multiple tables, let user select one
                if len(tables) > 1:
                    st.info(f"Found {len(tables)} tables in the PDF.")
                    table_idx = st.selectbox("Select a table to process:", 
                                            range(len(tables)), 
                                            format_func=lambda x: f"Table {x+1}")
                    st.session_state.df = tables[table_idx]
                else:
                    st.session_state.df = tables[0]
                
                st.success("PDF file loaded successfully!")
                
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        st.info("Make sure the PDF contains tables with clear structure.")

# Step 2: Display columns and selection
if st.session_state.df is not None:
    df = st.session_state.df
    
    st.header("Step 2: Select Columns")
    
    # Display preview of the data
    st.subheader("Data Preview")
    st.dataframe(df.head(10), use_container_width=True)
    
    st.subheader("Available Columns")
    columns = df.columns.tolist()
    
    # Option to select all or individual columns
    col1, col2 = st.columns([1, 3])
    
    with col1:
        select_all = st.checkbox("Select All Columns", value=True)
    
    with col2:
        if select_all:
            st.session_state.selected_columns = columns
            st.info(f"All {len(columns)} columns selected")
        else:
            st.session_state.selected_columns = st.multiselect(
                "Choose columns to include:",
                columns,
                default=st.session_state.selected_columns if st.session_state.selected_columns else columns
            )
    
    # Step 3: Reorder columns
    if st.session_state.selected_columns:
        st.header("Step 3: Reorder Columns")
        
        st.write("Drag and drop to reorder columns (or use the interface below):")
        
        # Create a simple reordering interface
        current_order = st.session_state.selected_columns.copy()
        
        reordered_columns = []
        for i, col in enumerate(current_order):
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.write(f"{i+1}. {col}")
            with col2:
                if i > 0:
                    if st.button("↑", key=f"up_{i}"):
                        current_order[i], current_order[i-1] = current_order[i-1], current_order[i]
                        st.session_state.selected_columns = current_order
                        st.rerun()
            with col3:
                if i < len(current_order) - 1:
                    if st.button("↓", key=f"down_{i}"):
                        current_order[i], current_order[i+1] = current_order[i+1], current_order[i]
                        st.session_state.selected_columns = current_order
                        st.rerun()
        
        st.session_state.column_order = current_order
        
        # Step 4: Preview and Download
        st.header("Step 4: Download Excel File")
        
        # Create final dataframe with selected and ordered columns
        final_df = df[st.session_state.column_order]
        
        st.subheader("Final Preview")
        st.dataframe(final_df.head(10), use_container_width=True)
        st.info(f"Total rows: {len(final_df)}, Total columns: {len(final_df.columns)}")
        
        # Convert to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Data')
        
        excel_data = output.getvalue()
        
        # Download button
        st.download_button(
            label="📥 Download Excel File",
            data=excel_data,
            file_name="converted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success("Your Excel file is ready for download!")
    else:
        st.warning("Please select at least one column to proceed.")

# Instructions
with st.sidebar:
    st.header("📖 Instructions")
    st.markdown("""
    1. **Upload PDF**: Click 'Browse files' and select your PDF file
    2. **Select Columns**: Choose which columns to include in the Excel file
    3. **Reorder Columns**: Use ↑ and ↓ buttons to change column order
    4. **Download**: Click the download button to get your Excel file
    
    ---
    
    **Requirements:**
    - PDF must contain tables
    - Tables should have clear column headers
    
    **Tip:** For best results, use PDFs with well-structured tables.
    """)
