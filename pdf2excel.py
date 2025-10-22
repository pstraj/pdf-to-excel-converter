import streamlit as st
import pandas as pd
import pdfplumber
import io

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
st.footer("Prashanth Rajashekar")
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    try:
        # Read PDF and extract tables
        with st.spinner("Reading PDF and extracting tables..."):
            pdf_bytes = uploaded_file.read()
            
            # Use pdfplumber to extract tables
            all_tables = []
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    # Extract tables from the page
                    page_tables = page.extract_tables()
                    
                    if page_tables:
                        for table_num, table in enumerate(page_tables, 1):
                            if table and len(table) > 1:  # Must have header and at least one data row
                                # First row is header
                                headers = table[0]
                                data_rows = table[1:]
                                
                                # Clean headers - remove None and empty strings
                                cleaned_headers = []
                                for i, h in enumerate(headers):
                                    if h and str(h).strip():
                                        cleaned_headers.append(str(h).strip())
                                    else:
                                        cleaned_headers.append(f"Column_{i+1}")
                                
                                # Create DataFrame
                                try:
                                    df_temp = pd.DataFrame(data_rows, columns=cleaned_headers)
                                    
                                    # Remove completely empty rows
                                    df_temp = df_temp.dropna(how='all')
                                    
                                    # Remove completely empty columns
                                    df_temp = df_temp.dropna(axis=1, how='all')
                                    
                                    if not df_temp.empty and len(df_temp.columns) > 0:
                                        all_tables.append({
                                            'df': df_temp,
                                            'page': page_num,
                                            'table_num': table_num,
                                            'rows': len(df_temp),
                                            'cols': len(df_temp.columns),
                                            'headers': list(df_temp.columns)
                                        })
                                except Exception as e:
                                    st.warning(f"Could not parse table {table_num} on page {page_num}: {str(e)}")
            
            if len(all_tables) == 0:
                st.error("❌ No tables found in the PDF file.")
                st.info("**Tips:**\n- Make sure your PDF contains tables with clear rows and columns\n- The PDF should not be a scanned image\n- Tables should have visible borders or clear structure")
            else:
                st.success(f"✅ Found {len(all_tables)} table(s) across {len(pdf.pages)} page(s)!")
                
                # Option to merge tables or select individual table
                if len(all_tables) > 1:
                    st.subheader("Table Processing Options")
                    
                    merge_option = st.radio(
                        "How would you like to process the tables?",
                        ["Merge all tables into one", "Select a specific table"],
                        index=0
                    )
                    
                    if merge_option == "Merge all tables into one":
                        # Check if all tables have the same columns
                        first_headers = set(all_tables[0]['headers'])
                        same_structure = all([set(t['headers']) == first_headers for t in all_tables])
                        
                        if same_structure:
                            # Merge all tables
                            merged_df = pd.concat([t['df'] for t in all_tables], ignore_index=True)
                            st.session_state.df = merged_df
                            
                            total_rows = sum(t['rows'] for t in all_tables)
                            st.info(f"✅ Merged {len(all_tables)} tables into one table with {total_rows} total rows")
                        else:
                            st.warning("⚠️ Tables have different column structures. Attempting to merge with all columns...")
                            
                            # Show column differences
                            with st.expander("View column differences"):
                                for i, t in enumerate(all_tables):
                                    st.write(f"**Page {t['page']}, Table {t['table_num']}:** {', '.join(t['headers'])}")
                            
                            # Merge with all columns (will have NaN for missing columns)
                            merged_df = pd.concat([t['df'] for t in all_tables], ignore_index=True, sort=False)
                            st.session_state.df = merged_df
                            
                            st.info(f"ℹ️ Merged {len(all_tables)} tables. Missing columns will show as empty cells.")
                    
                    else:
                        # Let user select a specific table
                        table_options = [
                            f"Page {t['page']}, Table {t['table_num']} ({t['rows']} rows × {t['cols']} columns)"
                            for t in all_tables
                        ]
                        
                        selected_table_idx = st.selectbox(
                            "Select a table to process:",
                            range(len(all_tables)),
                            format_func=lambda x: table_options[x]
                        )
                        
                        st.session_state.df = all_tables[selected_table_idx]['df']
                else:
                    # Only one table found
                    st.session_state.df = all_tables[0]['df']
                    st.info(f"Table found on page {all_tables[0]['page']} with {all_tables[0]['rows']} rows and {all_tables[0]['cols']} columns")
                
    except Exception as e:
        st.error(f"❌ Error reading PDF: {str(e)}")
        st.info("**Please ensure:**\n- The file is a valid PDF\n- The PDF contains tables (not just text or images)\n- The PDF is not password protected")

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
