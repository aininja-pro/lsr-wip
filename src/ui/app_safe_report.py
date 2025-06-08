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
sys.path.append('/app/src')

from data_processing.aggregation import (
    filter_gl_accounts, 
    compute_amounts, 
    aggregate_gl_data
)
from data_processing.merge_data import merge_wip_with_gl

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def initialize_session_state():
    """Initialize session state variables"""
    if 'files_uploaded' not in st.session_state:
        st.session_state.files_uploaded = {}
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'merged_data' not in st.session_state:
        st.session_state.merged_data = None

def map_columns_flexible(df, column_mapping):
    """Map column names flexibly using variations"""
    mapped_df = df.copy()
    
    for standard_name, variations in column_mapping.items():
        for variation in variations:
            if variation in df.columns:
                if variation != standard_name:
                    mapped_df = mapped_df.rename(columns={variation: standard_name})
                break
    
    return mapped_df

def process_data(wip_bytes, gl_bytes, include_closed):
    """Process the data using our existing functions"""
    try:
        with st.spinner("Processing GL data..."):
            # Load GL inquiry from bytes
            gl_df = pd.read_excel(io.BytesIO(gl_bytes))
            
            # Log available GL columns to help debug
            logger.info(f"Available GL Inquiry columns: {list(gl_df.columns)}")
            
            # Apply column mapping for GL data
            gl_column_variations = {
                'Account': ['Account', 'Account Number', 'Acct', 'GL Account'],
                'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number', 'Project No'],
                'Debit': ['Debit', 'Dr', 'Debit Amount'],
                'Credit': ['Credit', 'Cr', 'Credit Amount'],
                'Account Type': ['Account Type', 'Type', 'Category']
            }
            gl_df = map_columns_flexible(gl_df, gl_column_variations)
            
            # Process GL data step by step
            filtered_gl = filter_gl_accounts(gl_df)
            amounts_gl = compute_amounts(filtered_gl)
            gl_summary = aggregate_gl_data(amounts_gl)
            
            st.info(f"✅ Processed {len(gl_summary)} GL entries")
            
        with st.spinner("Merging data..."):
            # Load WIP worksheet from bytes
            wip_df = pd.read_excel(io.BytesIO(wip_bytes))
            
            # Log available columns to help debug
            logger.info(f"Available WIP Worksheet columns: {list(wip_df.columns)}")
            
            # Apply column mapping for WIP worksheet
            wip_column_variations = {
                'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number', 'Project No'],
                'Status': ['Status', 'Job Status', 'Project Status', 'State'],
                'Job Name': ['Job Name', 'Project Name', 'Description', 'Job Description'],
                'Budget Material': ['Budget Material', 'Material Budget', 'Mat Budget', 'Budget Mat'],
                'Budget Labor': ['Budget Labor', 'Labor Budget', 'Lab Budget', 'Budget Lab'],
                'Contract Amount': ['Contract Amount', 'Contract Value', 'Total Contract', 'Contract'],
                'Estimated Sub Labor': ['Estimated Sub Labor', 'Est Sub Labor', 'Sub Labor Budget', 'Sub Labor Est'],
                'Estimated Material': ['Estimated Material', 'Est Material', 'Material Budget', 'Material Est']
            }
            wip_df = map_columns_flexible(wip_df, wip_column_variations)
            
            # Log mapped columns
            logger.info(f"WIP Worksheet columns after mapping: {list(wip_df.columns)}")
            
            merged_df = merge_wip_with_gl(wip_df, gl_summary, include_closed)
            st.info(f"✅ Merged data for {len(merged_df)} jobs")
            
        return merged_df
        
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        logger.error(f"Processing error: {e}")
        return None

def generate_update_reports(merged_df):
    """Generate reports with EXACTLY the fields requested"""
    
    # 5040 Section - Labor Report (4 fields only)
    labor_data = []
    for _, job in merged_df.iterrows():
        labor_actual = job.get('5040', 0) or job.get('Labor Actual', 0) or job.get('Sub Labor', 0)
        
        labor_data.append({
            'Job Number': job.get('Job Number', ''),
            'Job Description': job.get('Job Name', job.get('Job Description', '')),
            'Contract Amount': job.get('Original Contract Amount', 0),  # Using actual column name
            'Estimated Sub Labor Costs': job.get('Total Subcontract Est', 0),  # Using actual column name
            'Monthly Sub Labor Costs': labor_actual,
            'Amount Billed': job.get('4020', 0)  # Using 4020 account data for billing
        })
    
    labor_df = pd.DataFrame(labor_data)
    
    # Convert to numeric and filter out rows where Monthly Sub Labor Costs is 0 or blank (include negative values)
    labor_df['Monthly Sub Labor Costs'] = pd.to_numeric(labor_df['Monthly Sub Labor Costs'], errors='coerce').fillna(0)
    labor_df = labor_df[labor_df['Monthly Sub Labor Costs'] != 0]
    
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

def create_excel_update_report(labor_df, material_df):
    """Create a comprehensive Excel report with all updates"""
    
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Labor section updates
        labor_df.to_excel(writer, sheet_name='5040_Labor_Updates', index=False)
        
        # Material section updates
        material_df.to_excel(writer, sheet_name='5030_Material_Updates', index=False)
        
        # Auto-adjust column widths for better readability
        for sheet_name in ['5040_Labor_Updates', '5030_Material_Updates']:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Add padding, max 50 chars
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Summary sheet
        summary_data = {
            'Section': ['5040 - Labor', '5030 - Material', 'Total'],
            'Jobs Count': [len(labor_df), len(material_df), len(labor_df)],
            'Total Contract Amount': [
                labor_df['Contract Amount'].sum(),
                0,  # Materials don't have contract amount
                labor_df['Contract Amount'].sum()
            ],
            'Total Actual': [
                labor_df['Monthly Sub Labor Costs'].sum(), 
                material_df['Monthly Material Costs'].sum(),
                labor_df['Monthly Sub Labor Costs'].sum() + material_df['Monthly Material Costs'].sum()
            ],
            'Total Budget': [
                labor_df['Estimated Sub Labor Costs'].sum(),
                material_df['Estimated Material Costs'].sum(), 
                labor_df['Estimated Sub Labor Costs'].sum() + material_df['Estimated Material Costs'].sum()
            ],
            'Total Variance': [
                labor_df['Monthly Sub Labor Costs'].sum() - labor_df['Estimated Sub Labor Costs'].sum(),
                material_df['Monthly Material Costs'].sum() - material_df['Estimated Material Costs'].sum(),
                (labor_df['Monthly Sub Labor Costs'].sum() - labor_df['Estimated Sub Labor Costs'].sum()) +
                (material_df['Monthly Material Costs'].sum() - material_df['Estimated Material Costs'].sum())
            ],
            'Total Amount Billed': [
                labor_df['Amount Billed'].sum(),
                0,  # Only labor section has amount billed
                labor_df['Amount Billed'].sum()
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Auto-adjust Summary sheet column widths
        summary_sheet = writer.sheets['Summary']
        for column in summary_sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            summary_sheet.column_dimensions[column_letter].width = adjusted_width
        
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
            "ADVANTAGES OF THIS APPROACH:",
            "✅ NO risk of corrupting your Excel file",
            "✅ ALL formulas and formatting preserved", 
            "✅ All macros and VBA code remain intact",
            "✅ You maintain full control over what gets updated",
            "✅ Easy to verify changes before applying them",
            "",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ]
        
        instructions_df = pd.DataFrame({'Instructions': instructions})
        instructions_df.to_excel(writer, sheet_name='Instructions', index=False)
        
        # Auto-adjust Instructions sheet column width
        instructions_sheet = writer.sheets['Instructions']
        instructions_sheet.column_dimensions['A'].width = 80  # Wide enough for instructions text
    
    buffer.seek(0)
    return buffer.getvalue()

def display_file_upload_section():
    """Display file upload interface"""
    st.markdown("#### 📁 File Upload")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Master WIP Report**")
        master_file = st.file_uploader(
            "Upload Master WIP Report",
            type=['xlsx', 'xlsm'],
            key='master_wip'
        )
        if master_file:
            st.success(f"✅ {master_file.name}")
    
    with col2:
        st.markdown("**WIP Worksheet Export**")
        wip_file = st.file_uploader(
            "Upload WIP Worksheet",
            type=['xlsx'],
            key='wip_worksheet'
        )
        if wip_file:
            st.session_state.files_uploaded['wip'] = wip_file.getvalue()
            st.success(f"✅ {wip_file.name}")
    
    with col3:
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
                    
                    # Display results
                    st.success("✅ Update reports generated successfully!")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.metric("📊 Total Jobs", len(merged_df))
                        st.metric("💼 Labor Actual", f"${labor_df['Monthly Sub Labor Costs'].sum():,.2f}")
                        st.metric("📦 Material Actual", f"${material_df['Monthly Material Costs'].sum():,.2f}")
                    
                    with col2:
                        labor_variance = labor_df['Monthly Sub Labor Costs'].sum() - labor_df['Estimated Sub Labor Costs'].sum()
                        material_variance = material_df['Monthly Material Costs'].sum() - material_df['Estimated Material Costs'].sum()
                        st.metric("📈 Labor Variance", f"${labor_variance:,.2f}")
                        st.metric("📈 Material Variance", f"${material_variance:,.2f}")
                        st.metric("📈 Total Variance", f"${labor_variance + material_variance:,.2f}")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Update Reports (Excel)",
                        data=excel_report,
                        file_name=f"WIP_Update_Reports_{month_year.replace(' ', '')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Download comprehensive update reports that you can use to manually update your WIP Excel file"
                    )
                    
                    # Preview data
                    st.markdown("#### 📋 Preview of Updates")
                    
                    tab1, tab2 = st.tabs(["5040 - Labor Updates", "5030 - Material Updates"])
                    
                    with tab1:
                        st.markdown("**Labor Section Updates**")
                        st.dataframe(labor_df, use_container_width=True)
                    
                    with tab2:
                        st.markdown("**Material Section Updates**") 
                        st.dataframe(material_df, use_container_width=True)
                    
            else:
                st.error("❌ Please upload at least the WIP Worksheet and GL Inquiry files")

if __name__ == "__main__":
    main() 