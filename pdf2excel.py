import streamlit as st
import pandas as pd
import io
from pypdf import PdfReader
import re

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
            pdf_bytes = uploaded_file.read()
            
            # Try to extract text and parse as CSV/table
            reader = PdfReader(io.BytesIO(pdf_bytes))
            
            # Extract text from all pages
            text_data = []
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    text_data.append(text)
            
            if not text_data:
                st.error("Could not extract text from PDF. The PDF might be image-based or encrypted.")
            else:
                # Combine all text
                full_text = "\n".join(text_data)
                
                # Try to parse as table (looking for tabular data)
                lines = full_text.split('\n')
                
                # Simple heuristic: find lines that look like table rows
                table_lines = [line.strip() for line in lines if line.strip()]
                
                if len(table_lines) < 2:
                    st.error("Could not find table structure in PDF.")
                else:
                    # Ask user to specify delimiter
                    st.info("Attempting to parse table from PDF text...")
                    
                    # Try common delimiters
                    delimiter = st.selectbox("If the preview looks incorrect, try a different delimiter:", 
                                            [None, ",", ";", "|", "\t", "  "], 
                                            format_func=lambda x: "Auto-detect" if x is None else f"'{x}'")
                    
                    if delimiter is None:
                        # Auto-detect: try to split by multiple spaces
                        header = re.split(r'\s{2,}', table_lines[0])
                        data_rows = [re.split(r'\s{2,}', line) for line in table_lines[1:]]
                    else:
                        header = table_lines[0].split(delimiter)
                        data_rows = [line.split(delimiter) for line in table_lines[1:]]
                    
                    # Clean up header
                    header = [h.strip() for h in header if h.strip()]
                    
                    # Filter rows that match header length
                    valid_rows = []
                    for row in data_rows:
                        row_clean = [cell.strip() for cell in row]
                        if len(row_clean) == len(header):
                            valid_rows.append(row_clean)
                    
                    if valid_rows:
                        df_temp = pd.DataFrame(valid_rows, columns=header)
                        st.session_state.df = df_temp
                        st.success(f"PDF file loaded successfully! Found {len(valid_rows)} rows.")
                    else:
                        st.error("Could not parse table structure. Please try a different delimiter or upload a CSV/Excel file instead.")
                
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        st.info("Tip: For best results, try converting your PDF to CSV or Excel first, or use a PDF with clear table structure.")

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
        
        st.write("Drag column names to reorder them (or use the dropdowns below):")
        
        # Show current order
        st.info(f"Current order: {' → '.join(st.session_state.selected_columns)}")
        
        # Initialize column order if not set
        if 'column_order' not in st.session_state or set(st.session_state.column_order) != set(st.session_state.selected_columns):
            st.session_state.column_order = st.session_state.selected_columns.copy()
        
        # Create a reordering interface using selectboxes
        reordered = []
        available_cols = st.session_state.selected_columns.copy()
        
        st.write("**Select columns in the desired order:**")
        
        for position in range(len(st.session_state.selected_columns)):
            col_label, col_selector = st.columns([1, 4])
            
            with col_label:
                st.write(f"**Position {position + 1}:**")
            
            with col_selector:
                # Default to maintaining current order if it exists
                if position < len(st.session_state.column_order):
                    default_col = st.session_state.column_order[position]
                    if default_col in available_cols:
                        default_idx = available_cols.index(default_col)
                    else:
                        default_idx = 0
                else:
                    default_idx = 0
                
                selected = st.selectbox(
                    f"pos_{position}",
                    available_cols,
                    index=default_idx,
                    key=f"col_order_{position}",
                    label_visibility="collapsed"
                )
                
                reordered.append(selected)
                available_cols.remove(selected)
        
        # Update the column order in session state
        st.session_state.column_order = reordered
        
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
