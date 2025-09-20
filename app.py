import streamlit as st
import pandas as pd

# Set the page configuration for the Streamlit app
st.set_page_config(page_title="Excel Comparator", layout="wide")

def compare_excel_files(ref_file, comp_file):
    """
    Compares two Excel files by sheet names and column headers.

    Args:
        ref_file (UploadedFile): The reference Excel file uploaded by the user.
        comp_file (UploadedFile): The comparison Excel file uploaded by the user.

    Returns:
        dict: A dictionary containing the comparison results.
    """
    results = {
        'sheet_summary': {},
        'column_differences': {}
    }

    try:
        # Use pandas ExcelFile to efficiently get sheet names
        ref_excel = pd.ExcelFile(ref_file)
        comp_excel = pd.ExcelFile(comp_file)

        ref_sheets = set(ref_excel.sheet_names)
        comp_sheets = set(comp_excel.sheet_names)

        # --- 1. Compare Sheet Names ---
        common_sheets = sorted(list(ref_sheets.intersection(comp_sheets)))
        sheets_only_in_ref = sorted(list(ref_sheets.difference(comp_sheets)))
        sheets_only_in_comp = sorted(list(comp_sheets.difference(ref_sheets)))

        results['sheet_summary'] = {
            'common': common_sheets,
            'only_in_ref': sheets_only_in_ref,
            'only_in_comp': sheets_only_in_comp
        }

        # --- 2. Compare Column Headers for Common Sheets ---
        for sheet in common_sheets:
            # Read only the header row for efficiency
            ref_df = pd.read_excel(ref_excel, sheet_name=sheet, nrows=0)
            comp_df = pd.read_excel(comp_excel, sheet_name=sheet, nrows=0)

            ref_cols = set(ref_df.columns)
            comp_cols = set(comp_df.columns)

            cols_only_in_ref = sorted(list(ref_cols.difference(comp_cols)))
            
            # This line specifically detects columns that are new in the comparison file
            cols_only_in_comp = sorted(list(comp_cols.difference(ref_cols)))

            # Only store results if there are differences
            if cols_only_in_ref or cols_only_in_comp:
                results['column_differences'][sheet] = {
                    'only_in_ref': cols_only_in_ref,
                    'only_in_comp': cols_only_in_comp
                }

    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None

    return results

# --- Streamlit UI ---

st.title("📊 Excel File Comparator")
st.markdown("Upload a **reference** Excel file and a **comparison** file to check for differences in sheet names and column headers.")

col1, col2 = st.columns(2)

with col1:
    st.header("Reference File")
    ref_file_input = st.file_uploader("Upload the master/template Excel file", type=['xlsx', 'xls'], key="ref")

with col2:
    st.header("Comparison File")
    comp_file_input = st.file_uploader("Upload the file to compare against the reference", type=['xlsx', 'xls'], key="comp")

if ref_file_input and comp_file_input:
    if st.button("🚀 Compare Files", type="primary"):
        
        # --- Perform Comparison ---
        comparison_results = compare_excel_files(ref_file_input, comp_file_input)

        if comparison_results:
            st.divider()
            st.header("Comparison Results")

            # --- Display Sheet Name Comparison ---
            st.subheader("Sheet Name Analysis")
            summary = comparison_results['sheet_summary']
            
            if not summary['only_in_ref'] and not summary['only_in_comp']:
                st.success("✅ All sheet names match perfectly between the two files.")
            else:
                if summary['only_in_ref']:
                    st.warning(f"Sheets found only in **{ref_file_input.name}** (Reference): `{summary['only_in_ref']}`")
                if summary['only_in_comp']:
                    st.warning(f"Sheets found only in **{comp_file_input.name}** (Comparison): `{summary['only_in_comp']}`")
            
            st.info(f"Common sheets found in both files: `{summary['common']}`")

            st.divider()

            # --- Display Column Header Comparison ---
            st.subheader("Column Header Analysis (for common sheets)")
            col_diffs = comparison_results['column_differences']
            common_sheets = summary['common']

            if not col_diffs:
                st.success("✅ All common sheets have matching column headers.")
            else:
                for sheet in common_sheets:
                    with st.expander(f"Differences in sheet: **'{sheet}'**", expanded=sheet in col_diffs):
                        if sheet in col_diffs:
                            diff = col_diffs[sheet]
                            if diff['only_in_ref']:
                                st.markdown(f"Columns only in **Reference**: `{diff['only_in_ref']}`")
                            if diff['only_in_comp']:
                                # This section displays the new columns found in the comparison file
                                st.markdown(f"Columns only in **Comparison**: `{diff['only_in_comp']}`")
                        else:
                            st.success("✅ Column headers match perfectly.")