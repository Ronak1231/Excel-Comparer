import streamlit as st
import pandas as pd
import io

# Set the page configuration for the Streamlit app
st.set_page_config(page_title="Advanced Excel Comparator", layout="wide")

def compare_excel_files(ref_file, comp_file, case_insensitive, primary_key_col, compare_by_position):
    """
    Compares two Excel files, including sheet names, column headers, column order, and row data.

    Args:
        ref_file (UploadedFile): The reference Excel file.
        comp_file (UploadedFile): The comparison Excel file.
        case_insensitive (bool): Whether to ignore case in column headers.
        primary_key_col (str): The name of the column to use as a primary key for row alignment.
        compare_by_position (bool): If True, compares the first sheet of each file regardless of name.

    Returns:
        dict: A dictionary containing detailed comparison results.
        None: If an error occurs.
    """
    results = {
        'sheet_summary': {},
        'column_differences': {},
        'column_order_differences': {},
        'data_differences': {}
    }

    try:
        ref_excel = pd.ExcelFile(ref_file)
        comp_excel = pd.ExcelFile(comp_file)

        ref_sheets_list = ref_excel.sheet_names
        comp_sheets_list = comp_excel.sheet_names
        
        sheets_to_process = [] # Will be a list of tuples: (ref_sheet_name, comp_sheet_name)

        if compare_by_position:
            if not ref_sheets_list or not comp_sheets_list:
                st.error("One or both of the uploaded files have no sheets to compare.")
                return None
            ref_s = ref_sheets_list[0]
            comp_s = comp_sheets_list[0]
            sheets_to_process.append((ref_s, comp_s))
            results['sheet_summary'] = {
                'mode': 'position',
                'comparison_pair': (ref_s, comp_s),
                'only_in_ref': ref_sheets_list[1:],
                'only_in_comp': comp_sheets_list[1:]
            }

        else: # Original name-based comparison
            ref_sheets_set = set(ref_sheets_list)
            comp_sheets_set = set(comp_sheets_list)
            common_sheets = sorted(list(ref_sheets_set.intersection(comp_sheets_set)))
            sheets_to_process = [(sheet, sheet) for sheet in common_sheets]

            results['sheet_summary'] = {
                'mode': 'name',
                'common': common_sheets,
                'only_in_ref': sorted(list(ref_sheets_set.difference(comp_sheets_set))),
                'only_in_comp': sorted(list(comp_sheets_set.difference(ref_sheets_set)))
            }

        for ref_sheet_name, comp_sheet_name in sheets_to_process:
            sheet_key = ref_sheet_name  # Use reference sheet name as the key for storing results

            ref_df = pd.read_excel(ref_excel, sheet_name=ref_sheet_name)
            comp_df = pd.read_excel(comp_excel, sheet_name=comp_sheet_name)

            # --- Column Header Preparation ---
            ref_cols_clean = [str(c).strip() for c in ref_df.columns]
            comp_cols_clean = [str(c).strip() for c in comp_df.columns]

            if case_insensitive:
                ref_cols_map = {c.lower(): c for c in ref_cols_clean}
                comp_cols_map = {c.lower(): c for c in comp_cols_clean}
                ref_cols_set = set(ref_cols_map.keys())
                comp_cols_set = set(comp_cols_map.keys())
            else:
                ref_cols_set = set(ref_cols_clean)
                comp_cols_set = set(comp_cols_clean)
            
            # --- 2a. Compare Column Presence ---
            cols_only_in_ref = sorted(list(ref_cols_set.difference(comp_cols_set)))
            cols_only_in_comp = sorted(list(comp_cols_set.difference(ref_cols_set)))

            if cols_only_in_ref or cols_only_in_comp:
                results['column_differences'][sheet_key] = {
                    'only_in_ref': cols_only_in_ref,
                    'only_in_comp': cols_only_in_comp
                }
            
            common_cols_set = ref_cols_set.intersection(comp_cols_set)

            # --- 2b. Compare Column Order (for common columns only) ---
            ref_common_ordered = [c for c in ref_cols_clean if (c.lower() if case_insensitive else c) in common_cols_set]
            comp_common_ordered = [c for c in comp_cols_clean if (c.lower() if case_insensitive else c) in common_cols_set]
            
            if ref_common_ordered != comp_common_ordered:
                results['column_order_differences'][sheet_key] = {
                    'ref_order': ref_common_ordered,
                    'comp_order': comp_common_ordered
                }

            # --- 3. Row Data Comparison ---
            common_cols = sorted(list(common_cols_set))
            
            if not common_cols: # Skip data comparison if no common columns
                continue

            if case_insensitive:
                ref_common_cols_orig = [ref_cols_map[c] for c in common_cols]
                comp_common_cols_orig = [comp_cols_map[c] for c in common_cols]
            else:
                ref_common_cols_orig = common_cols
                comp_common_cols_orig = common_cols
            
            ref_df_subset = ref_df[ref_common_cols_orig]
            comp_df_subset = comp_df[comp_common_cols_orig]
            
            comp_df_subset.columns = ref_common_cols_orig

            use_pk = False
            if primary_key_col:
                pk_col_clean = primary_key_col.strip()
                if case_insensitive:
                    pk_col_clean = pk_col_clean.lower()
                
                if pk_col_clean in common_cols:
                    use_pk = True
                    pk_orig = ref_cols_map[pk_col_clean] if case_insensitive else pk_col_clean
                    
                    if ref_df[pk_orig].duplicated().any() or comp_df[pk_orig].duplicated().any():
                         results['data_differences'][sheet_key] = {'status': 'error', 'message': f"Error: Duplicate values found in the primary key column '{pk_orig}'. Cannot perform data comparison."}
                         continue

                    ref_df_subset = ref_df_subset.set_index(pk_orig)
                    comp_df_subset = comp_df_subset.set_index(pk_orig)

            try:
                aligned_ref, aligned_comp = ref_df_subset.align(comp_df_subset, join='outer', axis=0)
                
                diff_df = aligned_ref.compare(aligned_comp, result_names=('Reference', 'Comparison'))
                
                new_rows = aligned_comp[aligned_ref.isnull().all(axis=1)]
                deleted_rows = aligned_ref[aligned_comp.isnull().all(axis=1)]

                if not diff_df.empty or not new_rows.empty or not deleted_rows.empty:
                    results['data_differences'][sheet_key] = {
                        'status': 'found',
                        'modified': diff_df,
                        'new': new_rows,
                        'deleted': deleted_rows,
                        'pk_used': primary_key_col if use_pk else 'Row Index'
                    }

            except Exception as e:
                 results['data_differences'][sheet_key] = {'status': 'error', 'message': f"Could not compare data. Ensure data types are consistent. Details: {e}"}

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
        return None

    return results

# --- Streamlit UI ---
st.title("📊 Advanced Excel File Comparator")
st.markdown("""
This tool compares two Excel files and highlights the differences in:
- **Sheet Names**: Finds sheets that are common, added, or removed.
- **Column Headers & Order**: For each common sheet, it finds columns that are added, removed, or out of order.
- **Row Data**: For each common sheet, it identifies new rows, deleted rows, and modified cells.
""")

st.sidebar.header("⚙️ Comparison Settings")
compare_by_position = st.sidebar.checkbox(
    "Compare by position (first sheet vs. first sheet)",
    value=True,
    help="If checked, ignores sheet names and compares the first sheet of the reference file with the first sheet of the comparison file."
)
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
            comparison_results = compare_excel_files(ref_file_input, comp_file_input, case_insensitive_checkbox, primary_key_input, compare_by_position)

        if comparison_results:
            st.divider()
            st.header("✅ Comparison Results")

            # --- 1. Display Sheet Name Comparison ---
            st.subheader("Sheet Name Analysis")
            summary = comparison_results['sheet_summary']
            
            if summary.get('mode') == 'position':
                ref_s, comp_s = summary['comparison_pair']
                st.info(f"Comparing by position: Reference sheet **'{ref_s}'** vs. Comparison sheet **'{comp_s}'**.")
                if summary['only_in_ref']:
                    st.warning(f"Other sheets found in Reference (`{ref_file_input.name}`): `{summary['only_in_ref']}`")
                if summary['only_in_comp']:
                    st.warning(f"Other sheets found in Comparison (`{comp_file_input.name}`): `{summary['only_in_comp']}`")
            else:
                if not summary['only_in_ref'] and not summary['only_in_comp']:
                    st.success("All sheet names match perfectly across files.")
                else:
                    if summary['only_in_ref']:
                        st.warning(f"Sheets found only in Reference (`{ref_file_input.name}`): `{summary['only_in_ref']}`")
                    if summary['only_in_comp']:
                        st.warning(f"Sheets found only in Comparison (`{comp_file_input.name}`): `{summary['only_in_comp']}`")
                
                st.info(f"Common sheets being compared by name: `{summary['common']}`")


            # --- 2, 3 & 4. Display Column, Order, and Data Comparison ---
            sheets_for_detailed_analysis = []
            if summary.get('mode') == 'position' and 'comparison_pair' in summary:
                sheets_for_detailed_analysis.append(summary['comparison_pair'])
            elif summary.get('mode') == 'name' and 'common' in summary:
                sheets_for_detailed_analysis = [(s, s) for s in summary['common']]

            if sheets_for_detailed_analysis:
                st.divider()
                st.subheader("Detailed Sheet-by-Sheet Analysis")

                for ref_sheet, comp_sheet in sheets_for_detailed_analysis:
                    sheet_key_for_results = ref_sheet
                    
                    col_diffs = comparison_results['column_differences'].get(sheet_key_for_results)
                    col_order_diffs = comparison_results['column_order_differences'].get(sheet_key_for_results)
                    data_diffs_info = comparison_results['data_differences'].get(sheet_key_for_results)
                    
                    has_col_diffs = col_diffs and (col_diffs['only_in_ref'] or col_diffs['only_in_comp'])
                    has_col_order_diffs = col_order_diffs is not None
                    has_data_diffs = data_diffs_info and data_diffs_info.get('status') == 'found'
                    has_error = data_diffs_info and data_diffs_info.get('status') == 'error'
                    
                    if ref_sheet == comp_sheet:
                        expander_base_title = f"sheet: **'{ref_sheet}'**"
                    else:
                        expander_base_title = f"Reference sheet **'{ref_sheet}'** vs. Comparison sheet **'{comp_sheet}'**"

                    if has_col_diffs or has_col_order_diffs or has_data_diffs or has_error:
                        expander_title = f"❗️ Differences found in {expander_base_title}"
                        expanded_state = True
                    else:
                        expander_title = f"✅ No differences found in {expander_base_title}"
                        expanded_state = False

                    with st.expander(expander_title, expanded=expanded_state):
                        # Column differences
                        st.markdown("---")
                        st.markdown("##### Column Headers")
                        
                        column_issue_found = False
                        if has_col_diffs:
                            column_issue_found = True
                            if col_diffs['only_in_ref']:
                                st.markdown(f"Columns only in **Reference**: `{col_diffs['only_in_ref']}`")
                            if col_diffs['only_in_comp']:
                                st.markdown(f"Columns only in **Comparison** (new columns): `{col_diffs['only_in_comp']}`")

                        if has_col_order_diffs:
                            column_issue_found = True
                            st.warning("Order of common columns does not match.")
                            c1, c2 = st.columns(2)
                            with c1:
                                st.markdown("**Reference Order (Common Columns)**")
                                st.json(col_order_diffs['ref_order'])
                            with c2:
                                st.markdown("**Comparison Order (Common Columns)**")
                                st.json(col_order_diffs['comp_order'])

                        if not column_issue_found:
                            st.success("Column headers and their order match perfectly.")

                        # Data differences
                        st.markdown("---")
                        st.markdown(f"##### Row Data")

                        if has_error:
                            st.error(data_diffs_info['message'])
                        elif has_data_diffs:
                            st.info(f"Comparison based on: **{data_diffs_info['pk_used']}**")
                            
                            if not data_diffs_info['modified'].empty:
                                st.markdown("###### Modified Cells")
                                st.dataframe(data_diffs_info['modified'])

                            if not data_diffs_info['new'].empty:
                                st.markdown("###### New Rows (in Comparison file)")
                                st.dataframe(data_diffs_info['new'])

                            if not data_diffs_info['deleted'].empty:
                                st.markdown("###### Deleted Rows (not in Comparison file)")
                                st.dataframe(data_diffs_info['deleted'])
                        elif col_diffs or col_order_diffs:
                            st.info("No data differences found in the common columns.")
                        else:
                            st.success("Row data matches perfectly.")

