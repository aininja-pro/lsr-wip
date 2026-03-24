#!/usr/bin/env python3
"""
Safe WIP Report Automation - Report Generation Only
Generates update reports that can be manually copied into Excel
This avoids ALL Excel corruption issues
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime
from pathlib import Path
import logging

# Import our data processing functions
import sys
sys.path.append(str(Path(__file__).parent.parent))

from data_processing.aggregation import (
    filter_gl_accounts,
    compute_amounts,
    aggregate_gl_data
)
from data_processing.merge_data import merge_wip_with_gl
from data_processing.column_mapping import map_dataframe_columns

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def initialize_session_state():
    """Initialize session state variables"""
    defaults = {
        'files_uploaded': {},
        'merged_data': None,
        'results_ready': False,
        'labor_df': None,
        'material_df': None,
        'excel_report': None,
        'month_year': None,
        'gl_entries': 0,
    }
    for key, default in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default

def process_data(wip_bytes, gl_bytes, include_closed):
    """Process the data using our existing functions"""
    try:
        with st.spinner("Processing GL data..."):
            gl_df = pd.read_excel(io.BytesIO(gl_bytes))
            logger.info(f"Available GL Inquiry columns: {list(gl_df.columns)}")
            gl_df = map_dataframe_columns(gl_df, 'gl_inquiry')

            filtered_gl = filter_gl_accounts(gl_df)
            amounts_gl = compute_amounts(filtered_gl)
            gl_summary = aggregate_gl_data(amounts_gl)
            st.session_state.gl_entries = len(gl_summary)

        with st.spinner("Merging data..."):
            wip_df = pd.read_excel(io.BytesIO(wip_bytes))
            logger.info(f"Available WIP Worksheet columns: {list(wip_df.columns)}")
            wip_df = map_dataframe_columns(wip_df, 'wip_worksheet')
            logger.info(f"WIP Worksheet columns after mapping: {list(wip_df.columns)}")

            merged_df = merge_wip_with_gl(wip_df, gl_summary, include_closed)

        return merged_df

    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        logger.error(f"Processing error: {e}")
        return None

def generate_update_reports(merged_df):
    """Generate reports with EXACTLY the fields requested"""
    
    # 5040 Section - Labor Report (with Percent Complete column)
    labor_data = []
    for _, job in merged_df.iterrows():
        labor_actual = job.get('5040', 0) or job.get('Labor Actual', 0) or job.get('Sub Labor', 0)
        estimated_labor = job.get('Total Subcontract Est', 0)
        
        # Calculate percent complete (avoid division by zero, cap at 100%)
        if estimated_labor > 0:
            percent_complete = min((labor_actual / estimated_labor) * 100, 100.0)
        else:
            percent_complete = 0.0
        
        labor_data.append({
            'Job Number': job.get('Job Number', ''),
            'Job Description': job.get('Job Name', job.get('Job Description', '')),
            'Contract Amount': job.get('Original Contract Amount', 0),  # Using actual column name
            'Estimated Sub Labor Costs': estimated_labor,  # Using actual column name
            'Monthly Sub Labor Costs': labor_actual,
            'Percent Complete': percent_complete,  # New column
            'Amount Billed': job.get('Amount Billed', 0)  # Using properly calculated Amount Billed from GL aggregation
        })
    
    labor_df = pd.DataFrame(labor_data)
    
    # Convert to numeric and filter to include jobs with labor costs OR billing
    labor_df['Monthly Sub Labor Costs'] = pd.to_numeric(labor_df['Monthly Sub Labor Costs'], errors='coerce').fillna(0)
    labor_df['Amount Billed'] = pd.to_numeric(labor_df['Amount Billed'], errors='coerce').fillna(0)
    
    # Include jobs that have either labor costs (!=0) OR have been billed
    labor_df = labor_df[(labor_df['Monthly Sub Labor Costs'] != 0) | (labor_df['Amount Billed'] > 0)]
    
    # 5030 Section - Material Report (4 fields only)
    material_data = []
    for _, job in merged_df.iterrows():
        material_actual = job.get('5030', 0) or job.get('Material Actual', 0) or job.get('Material', 0)
        
        material_data.append({
            'Job Number': job.get('Job Number', ''),
            'Job Description': job.get('Job Name', job.get('Job Description', '')),
            'Estimated Material Costs': job.get('Total Material Estimate', 0),  # Using actual column name
            'Monthly Material Costs': material_actual
        })
    
    material_df = pd.DataFrame(material_data)
    
    # Convert to numeric and filter out rows where Monthly Material Costs is 0 or blank (include negative values)
    material_df['Monthly Material Costs'] = pd.to_numeric(material_df['Monthly Material Costs'], errors='coerce').fillna(0)
    material_df = material_df[material_df['Monthly Material Costs'] != 0]
    
    return labor_df, material_df

def _auto_adjust_column_widths(worksheet):
    """Auto-adjust column widths based on cell content."""
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except (TypeError, AttributeError):
                pass
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)


def _apply_number_format(worksheet, col_names, fmt):
    """Apply a number format to named columns in a worksheet."""
    headers = [cell.value for cell in worksheet[1]]
    for col_name in col_names:
        if col_name in headers:
            col_index = headers.index(col_name) + 1
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=col_index)
                if cell.value is not None and isinstance(cell.value, (int, float)):
                    cell.number_format = fmt


def create_excel_update_report(labor_df, material_df):
    """Create a comprehensive Excel report with all updates"""

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        labor_df.to_excel(writer, sheet_name='5040_Labor_Updates', index=False)
        material_df.to_excel(writer, sheet_name='5030_Material_Updates', index=False)

        # Format data sheets
        currency_columns = {
            '5040_Labor_Updates': ['Contract Amount', 'Monthly Sub Labor Costs', 'Estimated Sub Labor Costs', 'Amount Billed'],
            '5030_Material_Updates': ['Monthly Material Costs', 'Estimated Material Costs']
        }

        for sheet_name in ['5040_Labor_Updates', '5030_Material_Updates']:
            worksheet = writer.sheets[sheet_name]
            _apply_number_format(worksheet, currency_columns[sheet_name], '$#,##0.00')

            # Percent Complete needs value conversion (already *100) before formatting
            if sheet_name == '5040_Labor_Updates':
                headers = [cell.value for cell in worksheet[1]]
                if 'Percent Complete' in headers:
                    col_index = headers.index('Percent Complete') + 1
                    for row in range(2, worksheet.max_row + 1):
                        cell = worksheet.cell(row=row, column=col_index)
                        if cell.value is not None and isinstance(cell.value, (int, float)):
                            cell.value = cell.value / 100
                            cell.number_format = '0.00%'

            _auto_adjust_column_widths(worksheet)

        # Summary sheet
        labor_actual_total = labor_df['Monthly Sub Labor Costs'].sum()
        material_actual_total = material_df['Monthly Material Costs'].sum()
        labor_budget_total = labor_df['Estimated Sub Labor Costs'].sum()
        material_budget_total = material_df['Estimated Material Costs'].sum()
        labor_variance = labor_actual_total - labor_budget_total
        material_variance = material_actual_total - material_budget_total
        contract_total = labor_df['Contract Amount'].sum()
        billed_total = labor_df['Amount Billed'].sum()

        summary_data = {
            'Section': ['5040 - Labor', '5030 - Material', 'Total'],
            'Jobs Count': [len(labor_df), len(material_df), len(labor_df)],
            'Total Contract Amount': [contract_total, 0, contract_total],
            'Total Actual': [labor_actual_total, material_actual_total, labor_actual_total + material_actual_total],
            'Total Budget': [labor_budget_total, material_budget_total, labor_budget_total + material_budget_total],
            'Total Variance': [labor_variance, material_variance, labor_variance + material_variance],
            'Total Amount Billed': [billed_total, 0, billed_total]
        }

        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        summary_sheet = writer.sheets['Summary']
        _apply_number_format(summary_sheet, ['Total Contract Amount', 'Total Actual', 'Total Budget', 'Total Variance', 'Total Amount Billed'], '$#,##0.00')
        _auto_adjust_column_widths(summary_sheet)

        # Instructions sheet
        instructions = [
            "WIP REPORT UPDATE INSTRUCTIONS",
            "",
            "This report contains all the updates for your WIP Report without modifying the original file.",
            "This approach preserves ALL formulas, formatting, and macros in your Excel file.",
            "",
            "HOW TO USE:",
            "",
            "1. LABOR SECTION (5040):",
            "   - Open the '5040_Labor_Updates' tab in this report",
            "   - Copy the 'Monthly Sub Labor Costs' column values",
            "   - Paste them into the appropriate column in your WIP Report's 5040 section",
            "",
            "2. MATERIAL SECTION (5030):",
            "   - Open the '5030_Material_Updates' tab in this report",
            "   - Copy the 'Monthly Material Costs' column values",
            "   - Paste them into the appropriate column in your WIP Report's 5030 section",
            "",
            "3. VERIFICATION:",
            "   - Check the 'Summary' tab for totals and variance analysis",
            "   - Variances > $1,000 should be reviewed",
            "",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ]

        instructions_df = pd.DataFrame({'Instructions': instructions})
        instructions_df.to_excel(writer, sheet_name='Instructions', index=False)
        writer.sheets['Instructions'].column_dimensions['A'].width = 80

    buffer.seek(0)
    return buffer.getvalue()

def display_file_upload_section():
    """Display file upload interface"""
    st.markdown("#### 📁 File Upload")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**WIP Worksheet Export**")
        wip_file = st.file_uploader(
            "Upload WIP Worksheet",
            type=['xlsx'],
            key='wip_worksheet'
        )
        if wip_file:
            st.session_state.files_uploaded['wip'] = wip_file.getvalue()
            st.success(f"✅ {wip_file.name}")
    
    with col2:
        st.markdown("**GL Inquiry Export**")
        gl_file = st.file_uploader(
            "Upload GL Inquiry",
            type=['xlsx'],
            key='gl_inquiry'
        )
        if gl_file:
            st.session_state.files_uploaded['gl'] = gl_file.getvalue()
            st.success(f"✅ {gl_file.name}")

def display_sidebar_options():
    """Display processing options in sidebar"""
    st.sidebar.markdown("### ⚙️ Options")
    
    # Month/Year selector with proper dropdown
    st.sidebar.markdown("**Report Period**")
    
    current_year = datetime.now().year
    years = list(range(current_year - 2, current_year + 2))
    months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ]
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        selected_month = st.selectbox("Month", months, index=3)  # Default to Apr
    with col2:
        selected_year = st.selectbox("Year", years, index=len(years)//2)
    
    # Format as MMM YY
    month_year = f"{selected_month} {str(selected_year)[-2:]}"
    
    st.sidebar.markdown("**Processing Settings**")
    include_closed = st.sidebar.checkbox(
        "Include Closed Jobs", 
        value=False,
        help="Check this to include jobs with 'Closed' status in the report. Useful for quarterly reviews."
    )
    
    return include_closed, month_year

def main():
    st.set_page_config(
        page_title="WIP Report Automation",
        page_icon="📊",
        layout="wide"
    )
    
    # Smaller, cleaner title
    st.markdown("# 📊 WIP Report Automation")
    st.markdown("*Generate update reports without modifying Excel files*")
    st.markdown("---")
    
    initialize_session_state()
    
    # Sidebar options
    include_closed, month_year = display_sidebar_options()
    
    # File Upload Section (main content)
    display_file_upload_section()
    
    st.markdown("---")
    
    # Process Button - better positioned and styled
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Generate Update Reports", type="primary", use_container_width=True):
            if len(st.session_state.files_uploaded) >= 2:  # Only need WIP and GL
                
                # Process the data
                merged_df = process_data(
                    st.session_state.files_uploaded['wip'],
                    st.session_state.files_uploaded['gl'],
                    include_closed
                )
                
                if merged_df is not None:
                    st.session_state.merged_data = merged_df
                    
                    # Generate reports
                    with st.spinner("Generating update reports..."):
                        labor_df, material_df = generate_update_reports(merged_df)
                        
                        # Create Excel report
                        excel_report = create_excel_update_report(labor_df, material_df)
                    
                    # Store results in session state for display
                    st.session_state.results_ready = True
                    st.session_state.labor_df = labor_df
                    st.session_state.material_df = material_df
                    st.session_state.excel_report = excel_report
                    st.session_state.month_year = month_year
                    
            else:
                st.error("❌ Please upload at least the WIP Worksheet and GL Inquiry files")
    
    # Download button - appears right after processing
    if st.session_state.get('results_ready', False):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.download_button(
                label="📥 Download Update Reports (Excel)",
                data=st.session_state.excel_report,
                file_name=f"WIP_Update_Reports_{st.session_state.month_year.replace(' ', '')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download comprehensive update reports that you can use to manually update your WIP Excel file",
                use_container_width=True
            )
    
    # Results Section - Same width as file upload section above
    if st.session_state.get('results_ready', False):
        st.markdown("---")
        
        # Create balanced full-width layout (same as file upload section)
        col1, col2 = st.columns([1, 1])
        
        with col1:
            # Processing Status Section
            st.markdown("### 🚀 Generate Update Reports")
            
            # Show processing status
            st.success(f"✅ Processed {st.session_state.gl_entries} GL entries")
            st.success(f"✅ Merged data for {len(st.session_state.merged_data)} jobs")
            st.success("✅ Update reports generated successfully!")
        
        with col2:
            # Combined Report Summary and Data Preview
            st.markdown("### 📊 Report Summary")
            
            # Calculate variances
            labor_variance = st.session_state.labor_df['Monthly Sub Labor Costs'].sum() - st.session_state.labor_df['Estimated Sub Labor Costs'].sum()
            material_variance = st.session_state.material_df['Monthly Material Costs'].sum() - st.session_state.material_df['Estimated Material Costs'].sum()
            total_variance = labor_variance + material_variance
            
            # Create a clean summary table
            summary_data = {
                'Category': ['Jobs Processed', 'Labor Actual', 'Material Actual', 'Labor Variance', 'Material Variance', 'Total Variance'],
                'Value': [
                    f"{len(st.session_state.merged_data)} jobs",
                    f"${st.session_state.labor_df['Monthly Sub Labor Costs'].sum():,.2f}",
                    f"${st.session_state.material_df['Monthly Material Costs'].sum():,.2f}",
                    f"${labor_variance:,.2f}",
                    f"${material_variance:,.2f}",
                    f"${total_variance:,.2f}"
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
        
        # Data Preview Section - Full width underneath
        st.markdown("---")
        st.markdown("### 📋 Data Preview")
        
        tab1, tab2 = st.tabs(["🔧 5040 - Labor Updates", "📦 5030 - Material Updates"])
        
        with tab1:
            st.markdown("**Labor Section Data (Non-Zero Values Only)**")
            if len(st.session_state.labor_df) > 0:
                st.dataframe(
                    st.session_state.labor_df, 
                    use_container_width=True,
                    height=450
                )
            else:
                st.info("No labor entries with non-zero values found.")
        
        with tab2:
            st.markdown("**Material Section Data (Non-Zero Values Only)**") 
            if len(st.session_state.material_df) > 0:
                st.dataframe(
                    st.session_state.material_df, 
                    use_container_width=True,
                    height=450
                )
            else:
                st.info("No material entries with non-zero values found.")

if __name__ == "__main__":
    main() 