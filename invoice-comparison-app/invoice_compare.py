import pandas as pd
import streamlit as st
import io
import re
import altair as alt

# --- 1. CUSTOM CSS FOR MODERN "PILL" TABS ---
def custom_tabs_style():
    """Applies a modern, segmented 'pill' style to the tabs."""
    st.markdown("""
        <style>
        /* 1. Style the Tab Container (The gray bar holding the tabs) */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px; /* Space between tabs */
            background-color: #F0F2F6; /* Light gray background track */
            padding: 8px; /* Padding around the tabs */
            border-radius: 10px; /* Rounded corners for the container */
            border: 1px solid #E0E0E0;
        }

        /* 2. Style the Inactive Tabs */
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: pre-wrap;
            background-color: transparent; /* Transparent background */
            border-radius: 6px; /* Rounded corners */
            color: #555555; /* Dark gray text */
            font-weight: 600; /* Semi-bold text */
            border: none; /* No border */
            flex-grow: 1; /* Make tabs expand to fill width */
            justify-content: center; /* Center text */
        }
        
        /* 3. Hover Effect for Inactive Tabs */
        .stTabs [data-baseweb="tab"]:hover {
            background-color: #E0E0E0; 
            color: #333333;
        }

        /* 4. Style the Active (Selected) Tab */
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            background-color: #581C87; /* ORLANDO PURPLE Background */
            color: #FFFFFF; /* White Text */
            box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1); /* Subtle shadow for depth */
        }
        
        /* 5. Hide the default red line that Streamlit puts under tabs */
        .stTabs [data-baseweb="tab-border"] {
            display: none;
        }
        
        /* 6. Add spacing below the tabs */
        .stTabs {
            margin-bottom: 20px;
        }
        </style>
    """, unsafe_allow_html=True)
# --- END CUSTOM CSS ---


# --- SESSION STATE INITIALIZATION ---
if 'current_filter' not in st.session_state:
    st.session_state['current_filter'] = None
if 'processed' not in st.session_state:
    st.session_state['processed'] = False
if 'amount_cols_to_process' not in st.session_state:
    st.session_state['amount_cols_to_process'] = []
if 'invoice_col' not in st.session_state:
    st.session_state['invoice_col'] = 'Invoice Number'
if 'selected_sheets' not in st.session_state:
    st.session_state['selected_sheets'] = []
# Ensure the state is always initialized
if 'show_advanced_settings' not in st.session_state:
    st.session_state['show_advanced_settings'] = False

# --- 2. CACHING FUNCTION TO SPEED UP FILE ACCESS ---
@st.cache_resource
def load_excel_metadata(uploaded_file):
    """Caches the ExcelFile object to avoid re-reading sheet names repeatedly."""
    excel_file = pd.ExcelFile(uploaded_file)
    return excel_file, excel_file.sheet_names

# --- 3. CALLBACK FUNCTION FOR CHECKBOX FIX ---
def toggle_advanced_settings():
    """A minimal callback function to ensure the checkbox state is updated immediately."""
    pass

# --- 4. OTHER HELPER FUNCTIONS ---
def color_summary_table(s):
    # Retained light colors for visual difference in downloaded Excel file only
    if s['MissingInSheets'] > 0:
        return ['background-color: #FFF3CD'] * len(s) # Light Yellow
    elif s['IsDuplicate']:
        return ['background-color: #F8D7DA'] * len(s) # Light Red
    elif s['AppearsInAllSheets']:
        return ['background-color: #D4EDDA'] * len(s) # Light Green
    else:
        return [''] * len(s)

def to_excel(df, sheet_name='Sheet1', engine='openpyxl'):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    if hasattr(df, 'data') and isinstance(df.data, pd.DataFrame): 
        df.data.to_excel(writer, index=False, sheet_name=sheet_name)
    else:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def filter_invoices(filter_type):
    df = st.session_state['final_summary']
    combined_df = st.session_state['combined']
    invoice_col_name = st.session_state.get('invoice_col', 'Invoice Number')

    if filter_type == 'all_sheets':
        filtered_df = df[df['AppearsInAllSheets'] == True].copy()
        title = "Invoices Available in ALL Sheets (Summary)"
    elif filter_type == 'missing':
        filtered_df = df[df['MissingInSheets'] > 0].copy()
        title = "Missing Invoices (in 1 or 2 sheets) (Summary)"
    elif filter_type == 'duplicates':
        duplicate_invoice_list = df[df['IsDuplicate'] == True][invoice_col_name].tolist()
        if duplicate_invoice_list:
            filtered_df = combined_df[combined_df[invoice_col_name].isin(duplicate_invoice_list)].copy()
            filtered_df.sort_values([invoice_col_name, 'Sheet'], inplace=True)
            if 'S. No.' in filtered_df.columns:
                 filtered_df.drop(columns=['S. No.'], inplace=True)
            filtered_df.insert(0, 'S. No.', range(1, 1 + len(filtered_df)))
            title = "CROSS-SHEET Duplicates (All Rows)"
        else:
            filtered_df = pd.DataFrame()
            title = "CROSS-SHEET Duplicates (All Rows)"
    elif filter_type == 'total':
        filtered_df = df.copy()
        title = "TOTAL Unique Invoices (Summary)"
    else:
        return None, None

    if 'S. No.' in filtered_df.columns and filter_type != 'duplicates':
        cols = ['S. No.'] + [col for col in filtered_df if col != 'S. No.']
        filtered_df = filtered_df[cols]

    return filtered_df, title


# --- CONFIGURATION ---
st.set_page_config(
    page_title="Invoice Reconciliation Dashboard 🎨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply the custom tab styles
custom_tabs_style() 


# --- SIDEBAR FOR INPUTS ---
uploaded_file = None
with st.sidebar:
    # Purple Orlando Header
    # st.markdown("<h1 style='color: #581C87; font-size: 28px; margin-bottom: 0px;'>🔮 Orlando</h1>", unsafe_allow_html=True)
    st.image("https://webtel.in/Images/webtel-logo.png", width=250)
    st.header("📄 File Uploader")
    st.markdown("---")

    uploaded_file = st.file_uploader("**Upload Excel file (.xlsx)**", type=["xlsx"])

    if uploaded_file:
        try:
            excel_file, sheet_names = load_excel_metadata(uploaded_file)
            st.success(f"✅ Sheets detected: {len(sheet_names)}")

            # Checkbox with callback fix
            st.checkbox(
                "⚙️ **Smart Filter**", 
                value=st.session_state['show_advanced_settings'],
                key='show_advanced_settings',
                on_change=toggle_advanced_settings
            )
            
            st.markdown("---")
            
            if st.session_state['show_advanced_settings']:
                st.markdown("##### 📝 Select Sheets")
                default_index_1 = 0 if len(sheet_names) > 0 else 0
                default_index_2 = 1 if len(sheet_names) > 1 else default_index_1
                sheet1_name = st.selectbox("Sheet 1 (Reference)", sheet_names, index=default_index_1, key="sheet_select_1")
                sheet2_name = st.selectbox("Sheet 2", sheet_names, index=default_index_2, key="sheet_select_2")
                
                sheet_names_with_blank = [""] + sheet_names
                default_index_3 = 0
                sheet3_name = st.selectbox("Sheet 3 (Optional)", sheet_names_with_blank, index=default_index_3, key="sheet_select_3")
                
                st.markdown("---")
                st.markdown("##### 🔑 Column Names")
                invoice_col = st.text_input("Invoice Column Name", "Invoice Number", key="input_invoice_col")
                st.session_state['invoice_col'] = invoice_col

                st.markdown("##### 💰 Columns for Amount/Value Difference (Max 3)")
                amount_col1 = st.text_input("Column 1 (e.g., Amount)", "TAX_STANDARD", key="input_amount_col1")
                amount_col2 = st.text_input("Column 2 (optional, e.g., Tax)", "", key="input_amount_col2")
                amount_col3 = st.text_input("Column 3 (optional, e.g., Total)", "", key="input_amount_col3")

                amount_cols_input = [col.strip() for col in [amount_col1, amount_col2, amount_col3] if col.strip()]
                st.session_state['amount_cols_to_process'] = amount_cols_input

            else:
                # Defaults
                if len(sheet_names) >= 2:
                    sheet1_name = sheet_names[0]
                    sheet2_name = sheet_names[1]
                    sheet3_name = "" 
                elif len(sheet_names) == 1:
                    sheet1_name = sheet_names[0]
                    sheet2_name = ""
                    sheet3_name = ""
                    st.warning(f"Only one sheet found ('{sheet1_name}'). Comparison requires two sheets.")
                else:
                    sheet1_name = ""
                    sheet2_name = ""
                    sheet3_name = ""
                    
                st.session_state['invoice_col'] = 'Invoice Number'
                st.session_state['amount_cols_to_process'] = ['TAX_STANDARD']
            
            st.markdown("---")
            
            invoice_col = st.session_state['invoice_col']
            amount_cols_input = st.session_state['amount_cols_to_process']

            if st.button("🚀 START INVOICE COMPARISON", type="primary", use_container_width=True):
                st.session_state['current_filter'] = None
                sheets_to_process = [sheet1_name, sheet2_name]
                if sheet3_name:
                    sheets_to_process.append(sheet3_name)
                
                sheets_to_process = [s for s in sheets_to_process if s]
                
                if len(sheets_to_process) < 2:
                    st.error("❌ At least two sheets must be available/selected for comparison.")
                    st.stop()

                st.session_state['selected_sheets'] = sheets_to_process
                total_sheets = len(sheets_to_process)

                try:
                    cols_to_read = [invoice_col] + amount_cols_input
                    sheet_dataframes = {}
                    for name in sheets_to_process:
                        sheet_df = pd.read_excel(
                            uploaded_file,
                            sheet_name=name,
                            dtype={invoice_col: str},
                            usecols=lambda x: x in cols_to_read if isinstance(x, str) else False
                        )
                        sheet_dataframes[name] = sheet_df

                    def clean_cols(df):
                        df.columns = [c.strip() for c in df.columns]
                        return df

                    all_cols_present = True
                    for name, df in sheet_dataframes.items():
                        sheet_dataframes[name] = clean_cols(df)
                        if invoice_col not in sheet_dataframes[name].columns:
                            all_cols_present = False
                            break

                    if not all_cols_present:
                        st.error(f"❌ Column '{invoice_col}' not found in all selected sheets.")
                    else:
                        def prepare(df, name):
                            cols_to_keep = [invoice_col] + [col for col in amount_cols_input if col in df.columns]
                            temp = df[cols_to_keep].dropna(subset=[invoice_col]).copy()
                            temp['Sheet'] = name
                            temp[invoice_col] = temp[invoice_col].astype(str).str.strip()
                            temp[invoice_col] = temp[invoice_col].str.replace(r'\.0$', '', regex=True)
                            temp[invoice_col] = temp[invoice_col].str.replace('\xa0', '').str.replace('\u200b', '')
                            temp = temp[temp[invoice_col].str.lower() != 'nan']
                            temp = temp[temp[invoice_col] != '']

                            for col in amount_cols_input:
                                if col in temp.columns:
                                    temp[col] = pd.to_numeric(temp[col], errors='coerce')
                            return temp

                        prepared_dfs = [prepare(df, name) for name, df in sheet_dataframes.items()]
                        combined = pd.concat(prepared_dfs, ignore_index=True)

                        summary = (
                            combined.groupby(invoice_col)
                            .agg(
                                SheetsAvailableIn=('Sheet', lambda x: ', '.join(sorted(set(x))), ),
                                TotalCount=('Sheet', 'count')
                            )
                            .reset_index()
                        )
                        summary['MissingInSheets'] = summary['TotalCount'].apply(lambda x: total_sheets - x if x < total_sheets else 0)
                        summary['IsDuplicate'] = summary['TotalCount'] > total_sheets
                        summary['AppearsInAllSheets'] = summary['TotalCount'] >= total_sheets
                        final_summary = summary

                        if amount_cols_input:
                            for col in amount_cols_input:
                                if col in combined.columns:
                                    pivot_amounts = combined.pivot_table(
                                        index=invoice_col,
                                        columns='Sheet',
                                        values=col,
                                        aggfunc='first'
                                    ).reset_index()

                                    diff_col_name = f"Difference_{col}"
                                    pivot_amounts[diff_col_name] = (
                                        pivot_amounts.drop(columns=[invoice_col])
                                        .apply(lambda x: x.max() - x.min() if x.count() > 1 else 0, axis=1)
                                    )

                                    new_cols = []
                                    for sheet_name_pivot in pivot_amounts.columns:
                                        if sheet_name_pivot in sheets_to_process:
                                            new_cols.append(f"{sheet_name_pivot.strip()}_{col}")
                                        else:
                                            new_cols.append(sheet_name_pivot)

                                    pivot_amounts.columns = new_cols

                                    final_summary = pd.merge(
                                        final_summary,
                                        pivot_amounts.drop(columns=[c for c in pivot_amounts.columns if c in sheets_to_process], errors='ignore'),
                                        on=invoice_col,
                                        how='outer'
                                    )

                        final_summary.insert(0, 'S. No.', range(1, 1 + len(final_summary)))

                        duplicates_within_sheets = (
                            combined.groupby(['Sheet', invoice_col])
                            .size()
                            .reset_index(name='Count')
                            .query('Count > 1')
                            .reset_index(drop=True)
                        )

                        if 'S. No.' in duplicates_within_sheets.columns:
                            duplicates_within_sheets.drop(columns=['S. No.'], inplace=True)

                        duplicates_within_sheets.insert(0, 'S. No.', range(1, 1 + len(duplicates_within_sheets)))
                        
                        st.session_state['final_summary'] = final_summary
                        st.session_state['combined'] = combined
                        st.session_state['duplicates'] = duplicates_within_sheets
                        st.session_state['processed'] = True
                        st.rerun()

                except Exception as e:
                    st.error(f"⚠️ Error during comparison logic: {e}")

        except Exception as e:
            st.error(f"❌ Could not read Excel file/sheets: {e}")

# --- MAIN CONTENT START ---
if st.session_state['processed']:
    final_summary = st.session_state['final_summary']
    duplicates_within_sheets = st.session_state['duplicates']
    combined = st.session_state['combined']

    st.header("Invoice Reconciliation Dashboard")

    st.subheader("💡 Key Invoice Statistics")
    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
    
    with col_stat1:
        st.metric(label="Total Unique Invoices", value=len(final_summary))
        if st.button("Total Records 📋", key='btn_total', use_container_width=True):
            st.session_state['current_filter'] = 'total'
    
    with col_stat2:
        count_all = len(final_summary[final_summary['AppearsInAllSheets'] == True])
        st.metric(label="In All Sheets", value=count_all, delta="Perfect Match 👍")
        if st.button("View All Sheets ✅", key='btn_all_sheets', use_container_width=True, type="primary"):
            st.session_state['current_filter'] = 'all_sheets'
            
    with col_stat3:
        count_missing = len(final_summary[final_summary['MissingInSheets'] > 0])
        st.metric(label="Missing Invoices", value=count_missing, delta="- Check Discrepancy", delta_color="inverse")
        if st.button("View Missing Invoices ⚠️", key='btn_missing', use_container_width=True, type="secondary"):
            st.session_state['current_filter'] = 'missing'

    with col_stat4:
        count_duplicates = len(final_summary[final_summary['IsDuplicate'] == True])
        st.metric(label="Cross-Sheet Duplicates", value=count_duplicates, delta="Investigation Needed 🚩", delta_color="off")
        if st.button("View Duplicates 🚩", key='btn_duplicates', use_container_width=True, type="secondary"):
            st.session_state['current_filter'] = 'duplicates'

    st.markdown("---")

    if st.session_state['current_filter']:
        st.subheader("🔍 Filtered Invoice List: Detailed View")
        filtered_df, title = filter_invoices(st.session_state['current_filter'])

        if filtered_df is not None and not filtered_df.empty:
            st.success(f"Showing {len(filtered_df)} rows for filter: **{title}**")
            st.data_editor(filtered_df, key=f"filtered_list_{st.session_state['current_filter']}", use_container_width=True, hide_index=True)
            
            col_down_filt, col_clear = st.columns([3, 1])
            with col_down_filt:
                excel_data = to_excel(filtered_df, sheet_name=title.replace(" ", "_").replace("(", "").replace(")", ""))
                st.download_button(label=f"📥 Download {title}", data=excel_data, file_name="Filtered_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
            with col_clear:
                if st.button("❌ Clear Filter", key='clear_filter_btn', use_container_width=True, type="secondary"):
                    st.session_state['current_filter'] = None
                    st.rerun()
        elif filtered_df is not None and filtered_df.empty:
              st.info(f"No invoices found for the filter: {title}")
        st.markdown("---")

    # --- TABS WITH CUSTOM STYLE ---
    tab_charts, tab_duplicates, tab_diff, tab_summary = st.tabs([
        "📊 Dashboard", 
        "🔁 Duplicates",  
        "💰 Amount Diff", 
        "📝 Full Summary" 
    ])

    with tab_charts:
        st.subheader("📊 Visual Distribution")
        chart_data = pd.DataFrame({
            'Status': ['In All Sheets', 'Missing in Some', 'Has Cross-Sheet Duplicates'],
            'Count': [
                len(final_summary[final_summary['AppearsInAllSheets'] == True]),
                len(final_summary[final_summary['MissingInSheets'] > 0]),
                len(final_summary[final_summary['IsDuplicate'] == True])
            ]
        })
        sheet_coverage = combined.groupby('Sheet')[st.session_state['invoice_col']].nunique().reset_index(name='Unique_Invoice_Count')
        chart_col1, chart_col2 = st.columns(2)
        with chart_col1:
            st.markdown("##### Invoice Status Breakdown")
            st.bar_chart(chart_data, x='Status', y='Count', color='Status', height=350)
        with chart_col2:
            st.markdown("##### Unique Invoices per Sheet")
            bar_chart = alt.Chart(sheet_coverage).mark_bar().encode(
                x=alt.X('Sheet', sort='-y'),
                y=alt.Y('Unique_Invoice_Count', title='Unique Invoice Count'),
                tooltip=['Sheet', 'Unique_Invoice_Count'],
                color=alt.Color('Sheet', scale=alt.Scale(range=['#581C87', '#9333EA', '#C084FC']))
            ).properties(title="Unique Invoices by Sheet")
            st.altair_chart(bar_chart, use_container_width=True)

        if st.session_state['amount_cols_to_process']:
            st.markdown("---")
            st.subheader("📉 Amount Differences Overview")
            diff_chart_data = []
            for col in st.session_state['amount_cols_to_process']:
                diff_col = f"Difference_{col}"
                if diff_col in final_summary.columns:
                    non_zero_diff_count = len(final_summary[final_summary[diff_col].abs() > 0.01])
                    diff_chart_data.append({'Column': col, 'Invoices with Difference': non_zero_diff_count})
            if diff_chart_data:
                df_diff_chart = pd.DataFrame(diff_chart_data)
                st.bar_chart(df_diff_chart, x='Column', y='Invoices with Difference', color='Column', height=350)

    with tab_duplicates:
        st.subheader("🔁 Duplicates Within Single Sheets")
        if len(duplicates_within_sheets):
            st.markdown("### 📋 Combined List of Internal Duplicates")
            st.data_editor(duplicates_within_sheets, key='internal_duplicates_table_combined', use_container_width=True, hide_index=True)
            duplicates_excel_data = to_excel(duplicates_within_sheets, sheet_name='Internal_Duplicates_List')
            st.download_button(label="📥 Download List (.xlsx)", data=duplicates_excel_data, file_name="Internal_Duplicates.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.markdown("---")
            st.markdown("### 📄 Duplicates Broken Down by Sheet")
            sheets_with_duplicates = duplicates_within_sheets['Sheet'].unique()
            for sheet_name in sheets_with_duplicates:
                sheet_duplicates_df = duplicates_within_sheets[duplicates_within_sheets['Sheet'] == sheet_name].copy()
                if 'S. No.' in sheet_duplicates_df.columns:
                    sheet_duplicates_df = sheet_duplicates_df.drop(columns=['S. No.'], errors='ignore')
                sheet_duplicates_df = sheet_duplicates_df.drop(columns=['Sheet'], errors='ignore')
                sheet_duplicates_df.insert(0, 'S. No.', range(1, 1 + len(sheet_duplicates_df)))
                st.markdown(f"#### Duplicates in **{sheet_name}** ({len(sheet_duplicates_df)} Invoices)")
                st.data_editor(sheet_duplicates_df, key=f'internal_duplicates_table_{sheet_name}_separate', use_container_width=True, hide_index=True)
                sheet_excel_data = to_excel(sheet_duplicates_df, sheet_name=f"{sheet_name}_Duplicates")
                st.download_button(label=f"📥 Download Duplicates in {sheet_name}", data=sheet_excel_data, file_name=f"Internal_Duplicates_{sheet_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"download_duplicates_{sheet_name}", use_container_width=True)
                st.markdown("---")
        else:
            st.info("No internal duplicates found.")

    with tab_diff:
        if st.session_state['amount_cols_to_process'] and len(st.session_state['selected_sheets']) >= 2:
            amount_col_name = st.session_state['amount_cols_to_process'][0]
            sheet1_name = st.session_state['selected_sheets'][0]
            sheet2_name = st.session_state['selected_sheets'][1]
            col_sheet1 = f"{sheet1_name}_{amount_col_name}"
            col_sheet2 = f"{sheet2_name}_{amount_col_name}"
            diff_col = f"Difference_{amount_col_name}"
            cols_to_select = [st.session_state['invoice_col'], col_sheet1, col_sheet2, diff_col]
            
            if all(col in final_summary.columns for col in cols_to_select):
                difference_df = final_summary[final_summary[diff_col].abs() > 0.01].copy()
                if not difference_df.empty:
                    final_diff_table = difference_df[[st.session_state['invoice_col'], col_sheet1, col_sheet2, diff_col]].copy()
                    if 'S. No.' in final_diff_table.columns: final_diff_table.drop(columns=['S. No.'], inplace=True)
                    final_diff_table.columns = ['Invoice Number', f'{sheet1_name} Amount', f'{sheet2_name} Amount', 'Difference in Amount']
                    final_diff_table.insert(0, 'S. No.', range(1, 1 + len(final_diff_table)))
                    st.subheader("💰 Invoices with Amount Difference (Focused)")
                    st.data_editor(final_diff_table, key='amount_difference_table_specific', use_container_width=True, hide_index=True)
                    diff_excel_data = to_excel(final_diff_table, sheet_name='Amount_Differences')
                    st.download_button(label="📥 Download Table (.xlsx)", data=diff_excel_data, file_name="Amount_Differences.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                else:
                    st.subheader("💰 Invoices with Amount Difference (Focused)")
                    st.info("🎉 Perfect Match! No invoices found with a non-zero difference in the primary amount column.")
            else:
                 st.subheader("💰 Invoices with Amount Difference (Focused)")
                 st.warning("Amount difference data is unavailable. Please ensure the primary amount column is present in the first two selected sheets.")
        else:
            st.subheader("💰 Invoices with Amount Difference (Focused)")
            st.info("To use this tab, please specify at least one amount column in the sidebar and select at least two sheets.")

    with tab_summary:
        st.subheader("📝 Full Comparison Summary")
        filtered_summary = final_summary.copy()
        if not filtered_summary.empty:
            st.info(f"Showing all **{len(final_summary)}** unique invoices.")
            st.dataframe(filtered_summary, key='full_summary_table_raw_fast', use_container_width=True, hide_index=True)
            unfiltered_data_summary = to_excel(filtered_summary.style.apply(color_summary_table, axis=1), sheet_name='Summary_Styled')
            st.download_button(label="Download Summary Data Only (Styled Excel) 📥", data=unfiltered_data_summary, file_name="Invoice_Summary_Table_Styled.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")
        else:
            st.warning("⚠️ No unique invoice data found.")

        st.markdown("---")
        st.subheader("📥 Download Complete Report (All Data)")
        col_down1, col_down2 = st.columns(2)
        with col_down1:
            unfiltered_data = to_excel(final_summary, sheet_name='Summary')
            st.download_button(label="Download Summary Data Only (Raw Excel) 📋", data=unfiltered_data, file_name="Invoice_Summary_Table_Raw.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")
        with col_down2:
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='openpyxl')
            styled_final_summary = final_summary.style.apply(color_summary_table, axis=1)
            styled_final_summary.data.to_excel(writer, sheet_name='Summary_Styled', index=False)
            duplicates_within_sheets.to_excel(writer, sheet_name='Internal_Duplicates', index=False)
            combined.to_excel(writer, sheet_name='Combined_Data', index=False)
            writer.close()
            complete_data = output.getvalue()
            st.download_button(label="Download All Data (3 Sheets in one .xlsx) 📦", data=complete_data, file_name="Invoice_Comparison_Report_Full.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

else:
    if not uploaded_file:
          st.info("⬆️ Please upload an Excel file and click 'START INVOICE COMPARISON' in the sidebar to begin processing.")
