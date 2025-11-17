import streamlit as st
import openpyxl
import pandas as pd
import io
# Import the ArrayFormula type to check against it
from openpyxl.worksheet.formula import ArrayFormula

st.set_page_config(layout="wide", page_title="Excel Formula Comparator")

# --- Helper Functions ---

def get_formulas(uploaded_file):
    """
    Reads an uploaded Excel file and extracts all formulas,
    returning a nested dictionary: {sheet_name: {cell_coord: formula_string}}
    """
    buffer = io.BytesIO(uploaded_file.getbuffer())
    # We must set read_only=False to properly load ArrayFormula objects
    wb = openpyxl.load_workbook(buffer, data_only=False, read_only=False, keep_links=False)
    
    all_formulas = {} 
    
    for sheet_name in wb.sheetnames:
        sheet_formulas = {}
        ws = wb[sheet_name]
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f':
                    formula_val = cell.value
                    
                    # --- FIX for ArrayFormula objects ---
                    # Check if it's an ArrayFormula object
                    if isinstance(formula_val, ArrayFormula):
                        # Get the text attribute from the object
                        sheet_formulas[cell.coordinate] = formula_val.text
                    else:
                        # Otherwise, just convert to string as before
                        sheet_formulas[cell.coordinate] = str(formula_val)
                    
        if sheet_formulas:
            all_formulas[sheet_name] = sheet_formulas
            
    return all_formulas, wb.sheetnames

def create_comparison_df(formulas_a, formulas_b, all_sheet_names):
    """
    Compares two formula dictionaries and returns a DataFrame
    of the differences.
    """
    comparison_data = [] 

    for sheet in all_sheet_names:
        sheet_a = formulas_a.get(sheet, {})
        sheet_b = formulas_b.get(sheet, {})
        
        all_cells_in_sheet = set(sheet_a.keys()) | set(sheet_b.keys())
        
        if not all_cells_in_sheet:
            continue 

        for cell in sorted(list(all_cells_in_sheet)):
            formula_a = sheet_a.get(cell) 
            formula_b = sheet_b.get(cell) 
            
            status = ""
            if formula_a == formula_b:
                status = "Identical"
            elif formula_a is None and formula_b is not None:
                status = "Only in File B"
            elif formula_a is not None and formula_b is None:
                status = "Only in File A"
            else: 
                status = "DIFFERENT"
                
            comparison_data.append({
                "Sheet": sheet,
                "Cell": cell,
                "File_A_Formula": formula_a,
                "File_B_Formula": formula_b,
                "Status": status
            })

    if not comparison_data:
        return pd.DataFrame(columns=["Sheet", "Cell", "File_A_Formula", "File_B_Formula", "Status"])
        
    return pd.DataFrame(comparison_data)

def style_diff_df(df):
    """
    Applies conditional styling to the DataFrame to highlight differences.
    """
    def style_row(row):
        style = [''] * len(row) # Default no style
        if row['Status'] == 'DIFFERENT':
            style = ['background-color: #FFF0F0'] * len(row) # Light red
        elif row['Status'] == 'Only in File A':
            style = ['background-color: #F0F8FF'] * len(row) # Light blue
        elif row['Status'] == 'Only in File B':
            style = ['background-color: #F0FFF0'] * len(row) # Light green
        elif row['Status'] == 'Identical':
            style = ['color: #999999'] * len(row) # Dim identical formulas
        return style

    return df.style.apply(style_row, axis=1)

# --- State Management Function ---
def clear_state():
    """
    Clears the comparison results from session state.
    This is called when a new file is uploaded.
    """
    if 'diff_df' in st.session_state:
        del st.session_state.diff_df
    if 'all_sheets' in st.session_state:
        del st.session_state.all_sheets

# --- Streamlit UI ---

st.title("Excel Formula Comparison Tool")
st.markdown("Upload two Excel files to see a cell-by-cell comparison of their formulas.")

col1, col2 = st.columns(2)

with col1:
    # Add on_change to clear results if a new file is uploaded
    file_a = st.file_uploader("Upload File A (e.g., Student 1)", type="xlsx", on_change=clear_state)

with col2:
    # Add on_change to clear results if a new file is uploaded
    file_b = st.file_uploader("Upload File B (e.g., Student 2)", type="xlsx", on_change=clear_state)

if file_a and file_b:
    
    # The button now *only* runs the analysis and saves to state
    if st.button("Compare Formulas"):
        
        with st.spinner("Analyzing formulas... This may take a moment."):
            # Get formula dictionaries and all unique sheet names
            formulas_a, sheets_a = get_formulas(file_a)
            formulas_b, sheets_b = get_formulas(file_b)
            all_sheets = sorted(list(set(sheets_a) | set(sheets_b)))
            
            # Create the comparison DataFrame
            diff_df = create_comparison_df(formulas_a, formulas_b, all_sheets)
            
            # --- SAVE TO SESSION STATE ---
            st.session_state.diff_df = diff_df
            st.session_state.all_sheets = all_sheets
            st.session_state.file_a_name = file_a.name
            st.session_state.file_b_name = file_b.name
        
        st.success("Comparison complete!")

# This block is now *outside* the button.
# It runs if the 'diff_df' exists in our session state.
if 'diff_df' in st.session_state:
    
    # Load data from session state
    diff_df = st.session_state.diff_df
    all_sheets = st.session_state.all_sheets
    
    # --- Display Summary ---
    st.subheader("Comparison Summary")
    stats_col1, stats_col2, stats_col3 = st.columns(3)
    
    total_identical = (diff_df['Status'] == 'Identical').sum()
    total_different = (diff_df['Status'] == 'DIFFERENT').sum()
    total_unique = ((diff_df['Status'] == 'Only in File A') | (diff_df['Status'] == 'Only in File B')).sum()
    
    stats_col1.metric("Identical Formulas", total_identical)
    stats_col2.metric("Different Formulas (Formulas present in both but different)", total_different)
    stats_col3.metric("Unique Formulas (formula in 1 but blank in another)", total_unique)
    
    # --- Display Filters ---
    st.subheader("Formula Comparison Table")
    
    # --- THIS IS THE CHANGED SECTION ---
    
    # 1. Filter by Status (Now a multiselect)
    # Get all possible statuses from the dataframe
    status_options = list(diff_df['Status'].unique())
    
    # Set default to show everything *except* "Identical"
    default_statuses = [s for s in status_options if s != 'Identical']
    
    # If there are no differences, default to showing "Identical"
    if not default_statuses and "Identical" in status_options:
        default_statuses = ["Identical"]
    
    selected_statuses = st.multiselect(
        "Filter by Status:", 
        options=status_options, 
        default=default_statuses # Default to showing differences
    )
    
    # 2. Filter by Sheet (Unchanged)
    sheet_options = ["All"] + all_sheets
    selected_sheet = st.selectbox("Filter by Sheet:", options=sheet_options, index=0)
    
    # --- Apply filters ---
    filtered_df = diff_df.copy()
    
    # Apply status filter (using .isin() for the list)
    if selected_statuses:
        filtered_df = filtered_df[filtered_df['Status'].isin(selected_statuses)]
    else:
        # If user deselects everything, show an empty table
        filtered_df = filtered_df.iloc[0:0] 
    
    # Apply sheet filter
    if selected_sheet != "All":
        filtered_df = filtered_df[filtered_df['Sheet'] == selected_sheet]

    # --- END OF CHANGED SECTION ---

    # Display the styled DataFrame
    st.dataframe(style_diff_df(filtered_df), height=600)

    # Provide a download button for the full diff
    csv = diff_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Full Comparison as CSV",
        data=csv,
        file_name=f"diff_{st.session_state.file_a_name}_vs_{st.session_state.file_b_name}.csv",
        mime="text/csv",
    )
elif file_a and file_b:
    st.info("Click 'Compare Formulas' to begin analysis.")
else:
    st.info("Please upload both files to compare.")