import streamlit as st
import os
import pandas as pd
import tempfile
from excel_manager import ExcelManager

st.title("Excel Manager App")

# Initialize session state
if 'excel_manager' not in st.session_state:
    st.session_state.excel_manager = None
if 'file_path' not in st.session_state:
    st.session_state.file_path = None
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = tempfile.mkdtemp()

# Function to reset the app
def reset_app():
    st.session_state.excel_manager = None
    st.session_state.file_path = None

# Sidebar for file operations
st.sidebar.header("File Operations")

# File upload
uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx", "xls"])
if uploaded_file is not None:
    # Save uploaded file to temp directory
    file_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Initialize ExcelManager with the uploaded file
    st.session_state.excel_manager = ExcelManager(file_path)
    st.session_state.file_path = file_path
    st.sidebar.success(f"Loaded: {uploaded_file.name}")

# Create new file
new_file_name = st.sidebar.text_input("Or create a new file (name.xlsx):")
if st.sidebar.button("Create New File") and new_file_name:
    if not new_file_name.endswith(('.xlsx', '.xls')):
        new_file_name += '.xlsx'
    
    file_path = os.path.join(st.session_state.temp_dir, new_file_name)
    st.session_state.excel_manager = ExcelManager()
    st.session_state.excel_manager.create_workbook(file_path)
    st.session_state.file_path = file_path
    st.sidebar.success(f"Created: {new_file_name}")

# Reset app
if st.sidebar.button("Reset"):
    reset_app()
    st.sidebar.success("Reset complete")

# Main content
if st.session_state.excel_manager is not None:
    st.subheader("Excel File Management")
    
    # Tabs for different operations
    tab1, tab2, tab3, tab4 = st.tabs(["Sheets", "Read", "Write", "Delete"])
    
    with tab1:
        st.subheader("Sheet Operations")
        
        # Count sheets
        if st.button("Count Sheets"):
            count = st.session_state.excel_manager.count_sheets()
            st.info(f"Number of sheets: {count}")
        
        # Get sheet names
        if st.button("Get Sheet Names"):
            names = st.session_state.excel_manager.get_sheet_names()
            st.info(f"Sheet names: {', '.join(names)}")
        
        # Create new sheet
        new_sheet_name = st.text_input("New sheet name:")
        if st.button("Create Sheet") and new_sheet_name:
            st.session_state.excel_manager.create_sheet(new_sheet_name)
            st.success(f"Created sheet: {new_sheet_name}")
            st.session_state.excel_manager.save()
    
    with tab2:
        st.subheader("Read Operations")
        
        # Select sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            selected_sheet = st.selectbox("Select sheet", sheet_names)
            
            # Read cell
            col1, col2, col3 = st.columns(3)
            with col1:
                read_row = st.number_input("Row", min_value=1, value=1, key="read_row")
            with col2:
                read_col = st.number_input("Column", min_value=1, value=1, key="read_col")
            with col3:
                resolve_references = st.checkbox("Resolve References", value=False, help="Follow cell references (like =A1) to get the actual value")
            
            if st.button("Read Cell"):
                value = st.session_state.excel_manager.read_cell(selected_sheet, read_row, read_col, hop=resolve_references)
                st.info(f"Cell value: {value}")
                
                # Display additional info if it's a formula
                if isinstance(value, str) and value.startswith('='):
                    st.info("This cell contains a formula reference. Check 'Resolve References' to see the referenced value.")
            
            # Read range
            st.subheader("Read Range")
            col1, col2 = st.columns(2)
            with col1:
                start_row = st.number_input("Start Row", min_value=1, value=1)
                end_row = st.number_input("End Row", min_value=1, value=5)
            with col2:
                start_col = st.number_input("Start Column", min_value=1, value=1)
                end_col = st.number_input("End Column", min_value=1, value=5)
            
            if st.button("Read Range"):
                values = st.session_state.excel_manager.read_range(
                    selected_sheet, start_row, start_col, end_row, end_col
                )
                # Convert to pandas DataFrame for better display
                df = pd.DataFrame(values)
                st.dataframe(df)
    
    with tab3:
        st.subheader("Write Operations")
        
        # Select sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            selected_sheet = st.selectbox("Select sheet", sheet_names, key="write_sheet")
            
            # Write cell
            col1, col2, col3 = st.columns(3)
            with col1:
                write_row = st.number_input("Row", min_value=1, value=1, key="write_row")
            with col2:
                write_col = st.number_input("Column", min_value=1, value=1, key="write_col")
            with col3:
                write_value = st.text_input("Value", key="write_value")
            
            if st.button("Write Cell"):
                st.session_state.excel_manager.write_cell(
                    selected_sheet, write_row, write_col, write_value
                )
                st.success(f"Wrote '{write_value}' to cell ({write_row}, {write_col})")
                st.session_state.excel_manager.save()
            
            # Write range (using CSV input)
            st.subheader("Write Range")
            col1, col2 = st.columns(2)
            with col1:
                start_row_write = st.number_input("Start Row", min_value=1, value=1, key="range_row")
            with col2:
                start_col_write = st.number_input("Start Column", min_value=1, value=1, key="range_col")
            
            csv_data = st.text_area(
                "Enter CSV data (comma-separated values, one row per line):",
                "1,2,3\n4,5,6\n7,8,9"
            )
            
            if st.button("Write Range"):
                # Parse CSV data
                rows = []
                for line in csv_data.strip().split("\n"):
                    values = line.split(",")
                    rows.append(values)
                
                st.session_state.excel_manager.write_range(
                    selected_sheet, start_row_write, start_col_write, rows
                )
                st.success(f"Wrote data to range starting at ({start_row_write}, {start_col_write})")
                st.session_state.excel_manager.save()
    
    with tab4:
        st.subheader("Delete Operations")
        
        # Delete sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            sheet_to_delete = st.selectbox("Select sheet to delete", sheet_names)
            
            if st.button("Delete Sheet") and len(sheet_names) > 1:
                st.session_state.excel_manager.delete_sheet(sheet_to_delete)
                st.success(f"Deleted sheet: {sheet_to_delete}")
                st.session_state.excel_manager.save()
            elif len(sheet_names) <= 1:
                st.error("Cannot delete the only sheet in the workbook.")
    
    # Download the file
    if st.session_state.file_path:
        with open(st.session_state.file_path, "rb") as file:
            file_name = os.path.basename(st.session_state.file_path)
            st.download_button(
                label="Download Excel file",
                data=file,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Please upload an Excel file or create a new one to start.")