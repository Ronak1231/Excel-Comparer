import streamlit as st
import pandas as pd
import io

# Set the page configuration for the Streamlit app
st.set_page_config(page_title="Advanced Excel Comparator", layout="wide")

def compare_excel_files(ref_file, comp_file, case_insensitive, primary_key_col):
    """
    Compares two Excel files, including sheet names, column headers, and row data.

    Args:
        ref_file (UploadedFile): The reference Excel file.
        comp_file (UploadedFile): The comparison Excel file.
        case_insensitive (bool): Whether to ignore case in column headers.
        primary_key_col (str): The name of the column to use as a primary key for row alignment.

    Returns:
        dict: A dictionary containing detailed comparison results.
        None: If an error occurs.
    """
    results = {
        'sheet_summary': {},
        'column_differences': {},
        'data_differences': {}
    }

    try:
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

        # --- 2 & 3. Compare Column Headers and Row Data for Common Sheets ---
        for sheet in common_sheets:
            ref_df = pd.read_excel(ref_excel, sheet_name=sheet)
            comp_df = pd.read_excel(comp_excel, sheet_name=sheet)

            # --- Column Header Comparison ---
            # Clean column names (strip whitespace, optionally convert to lower case)
            ref_cols_orig = ref_df.columns
            comp_cols_orig = comp_df.columns

            ref_cols_clean = [str(c).strip() for c in ref_cols_orig]
            comp_cols_clean = [str(c).strip() for c in comp_cols_orig]

            if case_insensitive:
                ref_cols_map = {c.lower(): c for c in ref_cols_clean}
                comp_cols_map = {c.lower(): c for c in comp_cols_clean}
                ref_cols_set = set(ref_cols_map.keys())
                comp_cols_set = set(comp_cols_map.keys())
            else:
                ref_cols_set = set(ref_cols_clean)
                comp_cols_set = set(comp_cols_clean)

            cols_only_in_ref = sorted(list(ref_cols_set.difference(comp_cols_set)))
            cols_only_in_comp = sorted(list(comp_cols_set.difference(ref_cols_set)))

            if cols_only_in_ref or cols_only_in_comp:
                results['column_differences'][sheet] = {
                    'only_in_ref': cols_only_in_ref,
                    'only_in_comp': cols_only_in_comp
                }

            # --- Row Data Comparison ---
            common_cols = sorted(list(ref_cols_set.intersection(comp_cols_set)))
            
            # Use original column names for indexing the dataframes
            if case_insensitive:
                ref_common_cols_orig = [ref_cols_map[c] for c in common_cols]
                comp_common_cols_orig = [comp_cols_map[c] for c in common_cols]
            else:
                ref_common_cols_orig = common_cols
                comp_common_cols_orig = common_cols

            ref_df_subset = ref_df[ref_common_cols_orig]
            comp_df_subset = comp_df[comp_common_cols_orig]
            
            # Rename comparison columns to match reference columns for comparison
            comp_df_subset.columns = ref_common_cols_orig

            # Align rows using primary key if provided
            use_pk = False
            if primary_key_col:
                pk_col_clean = primary_key_col.strip()
                if case_insensitive:
                    pk_col_clean = pk_col_clean.lower()
                
                if pk_col_clean in common_cols:
                    use_pk = True
                    # Use original case for the column name
                    pk_orig = ref_cols_map[pk_col_clean] if case_insensitive else pk_col_clean
                    
                    # Check for duplicate keys
                    if ref_df[pk_orig].duplicated().any() or comp_df[pk_orig].duplicated().any():
                         results['data_differences'][sheet] = {'status': 'error', 'message': f"Error: Duplicate values found in the primary key column '{pk_orig}'. Cannot perform data comparison."}
                         continue

                    ref_df_subset = ref_df_subset.set_index(pk_orig)
                    comp_df_subset = comp_df_subset.set_index(pk_orig)

            # Perform data comparison
            try:
                # The 'align' method is crucial to handle rows that might exist in one file but not the other
                aligned_ref, aligned_comp = ref_df_subset.align(comp_df_subset, join='outer', axis=0)
                
                # Using pandas compare, which is excellent for this task
                diff_df = aligned_ref.compare(aligned_comp, result_names=('Reference', 'Comparison'))
                
                # Find rows that are entirely new or deleted
                new_rows = aligned_comp[aligned_ref.isnull().all(axis=1)]
                deleted_rows = aligned_ref[aligned_comp.isnull().all(axis=1)]

                if not diff_df.empty or not new_rows.empty or not deleted_rows.empty:
                    results['data_differences'][sheet] = {
                        'status': 'found',
                        'modified': diff_df,
                        'new': new_rows,
                        'deleted': deleted_rows,
                        'pk_used': primary_key_col if use_pk else 'Row Index'
                    }

            except Exception as e:
                 results['data_differences'][sheet] = {'status': 'error', 'message': f"Could not compare data. Ensure data types are consistent. Details: {e}"}

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
        return None

    return results

# --- Streamlit UI ---
st.title("📊 Advanced Excel File Comparator")
st.markdown("""
This tool compares two Excel files and highlights the differences in:
- **Sheet Names**: Finds sheets that are common, added, or removed.
- **Column Headers**: For each common sheet, it finds columns that are added or removed.
- **Row Data**: For each common sheet, it identifies new rows, deleted rows, and modified cells.
""")

st.sidebar.header("⚙️ Comparison Settings")
case_insensitive_checkbox = st.sidebar.checkbox(
    "Ignore case in column names", 
    value=True,
    help="If checked, 'ColumnA' and 'columna' will be treated as the same column."
)
primary_key_input = st.sidebar.text_input(
    "Primary Key Column (optional)",
    help="Enter a column name with unique values (e.g., 'ID', 'ProductID'). This greatly improves the accuracy of row comparison."
)

col1, col2 = st.columns(2)

with col1:
    st.header("1. Reference File")
    ref_file_input = st.file_uploader("Upload the master or template file", type=['xlsx', 'xls'], key="ref")

with col2:
    st.header("2. Comparison File")
    comp_file_input = st.file_uploader("Upload the file to compare against the reference", type=['xlsx', 'xls'], key="comp")

if ref_file_input and comp_file_input:
    if st.button("🚀 Compare Files", type="primary", use_container_width=True):
        
        with st.spinner("Comparing files... this may take a moment."):
            comparison_results = compare_excel_files(ref_file_input, comp_file_input, case_insensitive_checkbox, primary_key_input)

        if comparison_results:
            st.divider()
            st.header("✅ Comparison Results")

            # --- 1. Display Sheet Name Comparison ---
            st.subheader("Sheet Name Analysis")
            summary = comparison_results['sheet_summary']
            
            if not summary['only_in_ref'] and not summary['only_in_comp']:
                st.success("Sheet names match perfectly between the two files.")
            else:
                if summary['only_in_ref']:
                    st.warning(f"Sheets found only in Reference (`{ref_file_input.name}`): `{summary['only_in_ref']}`")
                if summary['only_in_comp']:
                    st.warning(f"Sheets found only in Comparison (`{comp_file_input.name}`): `{summary['only_in_comp']}`")
            
            st.info(f"Common sheets found: `{summary['common']}`")

            # --- 2 & 3. Display Column and Data Comparison ---
            common_sheets = summary['common']
            if common_sheets:
                st.divider()
                st.subheader("Detailed Sheet-by-Sheet Analysis")

                for sheet in common_sheets:
                    col_diffs = comparison_results['column_differences'].get(sheet)
                    data_diffs_info = comparison_results['data_differences'].get(sheet)
                    
                    has_col_diffs = col_diffs and (col_diffs['only_in_ref'] or col_diffs['only_in_comp'])
                    has_data_diffs = data_diffs_info and data_diffs_info['status'] != 'error' and (not data_diffs_info['modified'].empty or not data_diffs_info['new'].empty or not data_diffs_info['deleted'].empty)
                    has_error = data_diffs_info and data_diffs_info['status'] == 'error'

                    # Determine expander title and state
                    if has_col_diffs or has_data_diffs or has_error:
                        expander_title = f"❗️ Differences found in sheet: **'{sheet}'**"
                        expanded_state = True
                    else:
                        expander_title = f"✅ No differences found in sheet: **'{sheet}'**"
                        expanded_state = False

                    with st.expander(expander_title, expanded=expanded_state):
                        # Column differences
                        st.markdown("---")
                        st.markdown("##### Column Headers")
                        if has_col_diffs:
                            if col_diffs['only_in_ref']:
                                st.markdown(f"Columns only in **Reference**: `{col_diffs['only_in_ref']}`")
                            if col_diffs['only_in_comp']:
                                st.markdown(f"Columns only in **Comparison** (new columns): `{col_diffs['only_in_comp']}`")
                        else:
                            st.success("Column headers match perfectly.")

                        # Data differences
                        st.markdown("---")
                        st.markdown(f"##### Row Data")

                        if has_error:
                            st.error(data_diffs_info['message'])
                        elif has_data_diffs:
                            st.info(f"Comparison based on: **{data_diffs_info['pk_used']}**")
                            
                            # Modified rows
                            if not data_diffs_info['modified'].empty:
                                st.markdown("###### Modified Cells")
                                st.dataframe(data_diffs_info['modified'])

                            # New rows
                            if not data_diffs_info['new'].empty:
                                st.markdown("###### New Rows (in Comparison file)")
                                st.dataframe(data_diffs_info['new'])

                            # Deleted rows
                            if not data_diffs_info['deleted'].empty:
                                st.markdown("###### Deleted Rows (not in Comparison file)")
                                st.dataframe(data_diffs_info['deleted'])
                        else:
                            st.success("Row data matches perfectly.")
