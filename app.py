import streamlit as st
import pandas as pd
import os
from datetime import datetime
import yaml
import plotly.express as px 
from st_aggrid import AgGrid, GridOptionsBuilder
from st_aggrid.grid_options_builder import GridOptionsBuilder

# Import your existing functions
from process_excel import (
    filter_sheets,
    process_data,
    update_stack_sheet,
    update_ar_sheet,
    load_and_process_data,
    generate_reports,
    cleanup_workbook
)

# Add at the start of your file, after imports:
st.set_page_config(
    page_title="MICU AR/AP Breakdown Analysis",
    layout="wide"  # This makes the app use the full width
)

def load_existing_data():
    try:
        # Load the existing analysis file
        file_path = 'summary_table_updated_analysis.xlsx'  # Adjust path as needed
        pc_overview_ap = pd.read_excel(file_path, sheet_name='pc_overview AP')
        pc_overview_ar = pd.read_excel(file_path, sheet_name='pc_overview AR')
        return pc_overview_ap, pc_overview_ar
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None, None

def create_aggrid_table(df, table_id):
    # Initialize session state for this table if not exists
    if f'filter_cleared_{table_id}' not in st.session_state:
        st.session_state[f'filter_cleared_{table_id}'] = False
    
    gb = GridOptionsBuilder.from_dataframe(df)
    
    # Configure each column
    gb.configure_default_column(
        filterable=True,
        sorteable=True,
        resizable=True,
        filter=True,
        floatingFilter=False,
        minWidth=150
    )
    
    # Enable filter icons and menu
    gb.configure_grid_options(
        enableRangeSelection=True,
        enableFilter=True,
        enableFilterToolPanel=True,
        domLayout='normal',
        showToolPanel=True,
        sideBar={
            'toolPanels': [
                {
                    'id': 'filters',
                    'labelDefault': 'Filters',
                    'labelKey': 'filters',
                    'iconKey': 'filter',
                    'toolPanel': 'agFiltersToolPanel',
                }
            ],
            'defaultToolPanel': ''
        },
        # Add function to clear filters and close panel
        onGridReady="""
        function(params) {
            window.gridApi_%s = params.api;
        }
        """ % table_id,
        onFirstDataRendered="""
        function(params) {
            window.clearFilters_%s = function() {
                window.gridApi_%s.setFilterModel(null);
                window.gridApi_%s.closeToolPanel();
            };
        }
        """ % (table_id, table_id, table_id)
    )
    
    grid_options = gb.build()
    
    # Add clear filters button and confirmation message
    col1, col2 = st.columns([6,1])
    with col2:
        if st.button('Clear All Filters', key=f'clear_filters_{table_id}'):
            st.session_state[f'filter_cleared_{table_id}'] = True
            # Add JavaScript to clear filters and close panel
            js_code = """
            <script>
                if (typeof clearFilters_%s === 'function') {
                    clearFilters_%s();
                }
            </script>
            """ % (table_id, table_id)
            st.components.v1.html(js_code, height=0)
            st.rerun()
    
    # Show confirmation message if filters were just cleared
    if st.session_state[f'filter_cleared_{table_id}']:
        st.success('Filters cleared successfully!')
        # Reset the state after showing the message
        st.session_state[f'filter_cleared_{table_id}'] = False
    
    return AgGrid(
        df,
        gridOptions=grid_options,
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True,
        update_mode='MODEL_CHANGED',
        theme='material',
        width='100%',
        height=500,
        custom_js=True
    )

def get_ap_column_descriptions():
    """Return descriptions for AP columns"""
    return {
        'PM Type': 'Primary classification (e.g., Chemical, UPW, WWT)',
        'Category': 'Used for categorizing Mapped_Category',
        'Mapped_Category': 'Standardized category mapping (e.g., Electrical, Mechanical, etc.)',
        'Scope': 'Project scope classification (e.g., Material, Labor, M+L)',
        'Project Number': 'Unique project identifier',
        'PO #': 'Purchase Order number',
        'Project Name': 'Full name of the project',
        'PO Description': 'Detailed description of the purchase order',
        'Vendor/Subcontractor': 'Name of the vendor or subcontractor',
        'Main/CO/DCR': 'Main Contract/Change Order/Design Change Request',
        'Actual Pertain': 'From All Stack, used for categorizing PM Type',
        'Type': 'From PC Overview, used for categorizing PM Type',
        'Type2': 'From PC Overview',
    }

def get_ar_column_descriptions():
    """Return descriptions for AR columns"""
    return {
        'Type': 'Category (e.g., Electrical, Mechanical, etc.)',
        'Project #': 'Unique project identifier',
        'Project Name': 'Full name of the project',
        'PO #': 'Purchase Order number',
        'Main Page': 'Primary classification (e.g., Chemical, UPW, WWT)',
        'CO/Added': 'Change Order or Additional work'
    }

def main():
    st.title("MICU AR/AP Breakdown Analysis")
    
    # Sidebar for configuration
    st.sidebar.header("Configuration")
    
    # Load config
    with open('config.yaml', 'r') as file:
        config = yaml.safe_load(file)
    
    # Main tabs - add AR Analysis
    tab1, tab2, tab3 = st.tabs(["AP Analysis", "AR Analysis", "Data View"])
    
    # AP Analysis Tab
    with tab1:
        st.header("AP Analysis")
        
        try:
            # Load the analysis file - AP sheet
            file_path = 'summary_table_updated_analysis.xlsx'
            df_ap = pd.read_excel(file_path, sheet_name='pc_overview AP')
            
            # Get AP column descriptions
            column_descriptions = get_ap_column_descriptions()
            
            # Get all possible columns for grouping
            groupable_columns = [col for col in df_ap.columns if col != 'Amount' 
                               and not pd.api.types.is_numeric_dtype(df_ap[col])]
            
            # Create formatted options for the multiselect
            formatted_options = []
            for col in groupable_columns:
                desc = column_descriptions.get(col, '')
                if desc:
                    formatted_options.append(f"{col} - {desc}")
                else:
                    formatted_options.append(col)
            
            # Find PM Type in formatted options
            default_option = next(
                (opt for opt in formatted_options if opt.startswith('PM Type')),
                formatted_options[0] if formatted_options else None
            )
            
            # Multi-select dropdown with descriptions - PM Type default for AP only
            selected_formatted = st.multiselect(
                'Select columns for analysis:',
                options=formatted_options,
                default=default_option,
                help="Hover over options to see descriptions",
                key='ap_select'
            )
            
            # Convert selected formatted options back to column names
            selected_columns = [opt.split(' - ')[0] for opt in selected_formatted]
            
            if selected_columns:
                # First, calculate total amount per PM Type
                pm_type_totals = df_ap.groupby('PM Type')['Amount'].sum().to_dict()
                
                # Create grouped analysis
                grouped_df = df_ap.groupby(selected_columns).agg({
                    'Amount': ['sum', 'count']
                }).reset_index()
                
                # Flatten column names and rename
                grouped_df.columns = [
                    col[0] if col[1] == '' else f"{col[0]}_{col[1]}" 
                    for col in grouped_df.columns
                ]
                grouped_df = grouped_df.rename(columns={
                    'Amount_sum': 'Total Amount',
                    'Amount_count': 'Count'
                })
                
                # Add percentage calculation if PM Type is selected
                if 'PM Type' in selected_columns:
                    # Calculate percentage
                    grouped_df['Percentage of PM Type'] = grouped_df.apply(
                        lambda row: (row['Total Amount'] / pm_type_totals[row['PM Type']]) * 100,
                        axis=1
                    )
                    
                    # Format percentage to 2 decimal places
                    grouped_df['Percentage of PM Type'] = grouped_df['Percentage of PM Type'].round(2)
                    
                    # Add % symbol
                    grouped_df['Percentage of PM Type'] = grouped_df['Percentage of PM Type'].astype(str) + '%'
                
                # Sort by PM Type and Total Amount descending
                grouped_df = grouped_df.sort_values(['PM Type', 'Total Amount'], ascending=[True, False])
                
                # Display results
                st.subheader("Breakdown By Selected Columns")
                
                # Show table
                st.dataframe(
                    grouped_df,
                    use_container_width=True,
                    hide_index=True
                )
                
                # Create visualization
                if len(selected_columns) <= 2:  # Bar chart for 1-2 columns
                    fig = px.bar(
                        grouped_df.head(10),  # Top 10 for readability
                        x=selected_columns[0],
                        y='Total Amount',
                        color=selected_columns[1] if len(selected_columns) > 1 else None,
                        title=f'Top 10 by Total Amount',
                        labels={'Total Amount': 'Total Amount ($)'},
                    )
                else:  # Treemap for 3+ columns
                    fig = px.treemap(
                        grouped_df,
                        path=selected_columns,
                        values='Total Amount',
                        title='Distribution of Total Amount'
                    )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # First clean up blank/NaN/0 values in Main/CO/DCR
                df_ap['Main/CO/DCR'] = df_ap['Main/CO/DCR'].fillna('Unspecified')
                df_ap['Main/CO/DCR'] = df_ap['Main/CO/DCR'].replace(['', '0', 0], 'Unspecified')

                # Define column order
                column_order = [
                    'Main Contract Scope',
                    'CO Scope (adding/additional scope)',
                    'DCR Scope',
                    'The budget execution does not pertain to this project',
                    'Unspecified'
                ]

                # Create pivot tables for percentages and amounts
                pivot_df = df_ap.pivot_table(
                    values='Amount',
                    index='PM Type',
                    columns='Main/CO/DCR',
                    aggfunc='sum'
                ).fillna(0)

                # Reorder columns (and add any missing columns with zeros)
                for col in column_order:
                    if col not in pivot_df.columns:
                        pivot_df[col] = 0
                pivot_df = pivot_df[column_order]

                # Calculate percentages
                pivot_pct = pivot_df.div(pivot_df.sum(axis=1), axis=0) * 100
                pivot_pct_display = pivot_pct.round(2)

                # Format with % symbol
                for column in pivot_pct_display.columns:
                    pivot_pct_display[column] = pivot_pct_display[column].astype(str) + '%'

                # Display the tables
                st.subheader("Percentage Breakdown by PM Type")
                st.dataframe(
                    pivot_pct_display,
                    use_container_width=True,
                    hide_index=False
                )

                # Also show the actual amounts
                st.subheader("Amount Breakdown by PM Type (in dollars)")
                st.dataframe(
                    pivot_df.round(2),
                    use_container_width=True,
                    hide_index=False
                )
                
                # Add download button for the analysis
                csv = grouped_df.to_csv(index=False)
                st.download_button(
                    "Download Analysis",
                    csv,
                    "analysis_results.csv",
                    "text/csv"
                )
                
        except Exception as e:
            st.error(f"Error in AP analysis: {str(e)}")
    
    # AR Analysis Tab
    with tab2:
        st.header("AR Analysis")
        
        try:
            # Load the analysis file - AR sheet
            file_path = 'summary_table_updated_analysis.xlsx'
            df_ar = pd.read_excel(file_path, sheet_name='pc_overview AR')
            
            # Get AR column descriptions
            column_descriptions = get_ar_column_descriptions()
            
            # Get all possible columns for grouping (excluding amount column)
            groupable_columns = [col for col in df_ar.columns if col != 'Total Contract $' 
                               and not pd.api.types.is_numeric_dtype(df_ar[col])]
            
            # Create formatted options for the multiselect
            formatted_options = []
            for col in groupable_columns:
                desc = column_descriptions.get(col, '')
                if desc:
                    formatted_options.append(f"{col} - {desc}")
                else:
                    formatted_options.append(col)
            
            # Find Main Page in formatted options
            default_option = next(
                (opt for opt in formatted_options if opt.startswith('Main Page')),
                formatted_options[0] if formatted_options else None
            )
            
            # Multi-select dropdown with descriptions - Main Page default for AR
            selected_formatted = st.multiselect(
                'Select columns for analysis:',
                options=formatted_options,
                default=default_option,
                help="Hover over options to see descriptions",
                key='ar_select'
            )
            
            # Convert selected formatted options back to column names
            selected_columns = [opt.split(' - ')[0] for opt in selected_formatted]
            
            if selected_columns:
                # Create grouped analysis using 'Total Contract $' instead of 'Amount'
                grouped_df = df_ar.groupby(selected_columns).agg({
                    'Total Contract $': ['sum', 'count']
                }).reset_index()
                
                # Flatten column names and rename
                grouped_df.columns = [
                    col[0] if col[1] == '' else f"{col[0]}_{col[1]}" 
                    for col in grouped_df.columns
                ]
                grouped_df = grouped_df.rename(columns={
                    'Total Contract $_sum': 'Total Amount',
                    'Total Contract $_count': 'Count'
                })
                
                # Sort by Total Amount descending
                grouped_df = grouped_df.sort_values('Total Amount', ascending=False)
                
                # Display results
                st.subheader("Breakdown By Selected Columns")
                
                # Show table
                st.dataframe(
                    grouped_df,
                    use_container_width=True,
                    hide_index=True
                )
                
                # Create visualization
                if len(selected_columns) <= 2:  # Bar chart for 1-2 columns
                    fig = px.bar(
                        grouped_df.head(10),  # Top 10 for readability
                        x=selected_columns[0],
                        y='Total Amount',
                        color=selected_columns[1] if len(selected_columns) > 1 else None,
                        title=f'Top 10 by Total Amount',
                        labels={'Total Amount': 'Total Amount ($)'},
                    )
                else:  # Treemap for 3+ columns
                    fig = px.treemap(
                        grouped_df,
                        path=selected_columns,
                        values='Total Amount',
                        title='Distribution of Total Amount'
                    )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # First clean up blank/NaN/0 values in CO/Added
                df_ar['CO/Added'] = df_ar['CO/Added'].fillna('Unspecified')
                df_ar['CO/Added'] = df_ar['CO/Added'].replace(['', '0', 0], 'Unspecified')

                # Define column order
                column_order = [
                    'Main',
                    'CO',
                    'Added',
                    'Unspecified'
                ]

                # Create pivot tables for percentages and amounts
                pivot_df = df_ar.pivot_table(
                    values='Total Contract $',
                    index='Main Page',
                    columns='CO/Added',
                    aggfunc='sum'
                ).fillna(0)

                # Reorder columns (and add any missing columns with zeros)
                for col in column_order:
                    if col not in pivot_df.columns:
                        pivot_df[col] = 0
                pivot_df = pivot_df[column_order]

                # Calculate percentages
                pivot_pct = pivot_df.div(pivot_df.sum(axis=1), axis=0) * 100
                pivot_pct_display = pivot_pct.round(2)

                # Format with % symbol
                for column in pivot_pct_display.columns:
                    pivot_pct_display[column] = pivot_pct_display[column].astype(str) + '%'

                # Display the tables
                st.subheader("Percentage Breakdown by Main Page")
                st.dataframe(
                    pivot_pct_display,
                    use_container_width=True,
                    hide_index=False
                )

                # Also show the actual amounts
                st.subheader("Amount Breakdown by Main Page (in dollars)")
                st.dataframe(
                    pivot_df.round(2),
                    use_container_width=True,
                    hide_index=False
                )
                
                # Add download button for the analysis
                csv = grouped_df.to_csv(index=False)
                st.download_button(
                    "Download Analysis",
                    csv,
                    "ar_analysis_results.csv",  # Different filename for AR
                    "text/csv",
                    key='ar_download'  # Unique key for AR
                )
                
        except Exception as e:
            st.error(f"Error in AR analysis: {str(e)}")
    
    # Data View Tab
    with tab3:
        st.header("Data Overview")
        pc_overview_ap, pc_overview_ar = load_existing_data()
        
        if pc_overview_ap is not None and pc_overview_ar is not None:
            # AP Overview Section
            st.subheader("PC Overview AP")
            ap_grid = create_aggrid_table(pc_overview_ap, 'ap')
            
            # AR Overview Section
            st.subheader("PC Overview AR")
            ar_grid = create_aggrid_table(pc_overview_ar, 'ar')

if __name__ == "__main__":
    main()