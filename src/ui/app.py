import streamlit as st
import pandas as pd
import sys
import os
from datetime import datetime
from pathlib import Path
import io
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import logging

# Add the src directory to the path so we can import our modules
sys.path.append(str(Path(__file__).parent.parent))

from data_processing.aggregation import aggregate_gl_data
from data_processing.merge_data import process_wip_merge, compute_variances
from data_processing.column_mapping import map_dataframe_columns
from data_processing.excel_integration_v2 import (
    load_wip_workbook, find_or_create_monthly_tab, find_section_markers,
    create_backup, update_wip_report_v2
)

# Configure Streamlit page
st.set_page_config(
    page_title="WIP Report Automation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 600;
        color: #2F80ED;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: 500;
        color: #333333;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #D4EDDA;
        border: 1px solid #C3E6CB;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #FFF3CD;
        border: 1px solid #FFEAA7;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #F8D7DA;
        border: 1px solid #F5C6CB;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    """Initialize session state variables"""
    if 'files_uploaded' not in st.session_state:
        st.session_state.files_uploaded = {
            'gl_inquiry': None,
            'wip_worksheet': None,
            'master_report': None
        }
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'backup_created' not in st.session_state:
        st.session_state.backup_created = None

def display_file_upload_section():
    """Display file upload widgets and validation"""
    st.markdown('<div class="section-header">📁 Upload Required Files</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("GL Inquiry Export")
        gl_file = st.file_uploader(
            "Choose GL Inquiry file",
            type=['xlsx', 'xls'],
            key="gl_inquiry_upload",
            help="Export from Sage containing actual costs (accounts 5040, 5030, 4020)"
        )
        if gl_file:
            st.session_state.files_uploaded['gl_inquiry'] = gl_file
            st.success("✓ GL Inquiry uploaded")
    
    with col2:
        st.subheader("WIP Worksheet Export")
        wip_file = st.file_uploader(
            "Choose WIP Worksheet file",
            type=['xlsx', 'xls'],
            key="wip_worksheet_upload",
            help="Export from Sage containing job budgets and status"
        )
        if wip_file:
            st.session_state.files_uploaded['wip_worksheet'] = wip_file
            st.success("✓ WIP Worksheet uploaded")
    
    with col3:
        st.subheader("Master WIP Report")
        master_file = st.file_uploader(
            "Choose Master WIP Report file",
            type=['xlsx', 'xlsm'],
            key="master_report_upload",
            help="The Excel file to be updated (will be backed up first)"
        )
        if master_file:
            st.session_state.files_uploaded['master_report'] = master_file
            st.success("✓ Master Report uploaded")
    
    # Show upload status
    files_ready = all(st.session_state.files_uploaded.values())
    if files_ready:
        st.markdown('<div class="success-box">✅ All required files uploaded and ready for processing!</div>', 
                   unsafe_allow_html=True)
    else:
        missing_files = [name for name, file in st.session_state.files_uploaded.items() if file is None]
        st.markdown(f'<div class="warning-box">⚠️ Still need: {", ".join(missing_files)}</div>', 
                   unsafe_allow_html=True)
    
    return files_ready

def display_processing_options():
    """Display processing configuration options"""
    st.markdown('<div class="section-header">⚙️ Processing Options</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Month/Year selector
        current_date = datetime.now()
        selected_date = st.date_input(
            "Select Month/Year to Process",
            value=current_date,
            help="Choose the month and year for the WIP report tab"
        )
        month_year = selected_date.strftime("%b %y")  # Format: "Jun 25"
        
        # Include closed jobs option
        include_closed = st.checkbox(
            "Include Closed Jobs",
            value=False,
            help="Check to include jobs with Status='Closed' (useful for quarterly true-ups)"
        )
    
    with col2:
        # Preview option
        preview_only = st.checkbox(
            "Preview Only (Don't Update Files)",
            value=True,
            help="Check to preview changes without updating the Master WIP Report"
        )
        
        # Backup option
        create_backup = st.checkbox(
            "Create Backup Before Update",
            value=True,
            help="Recommended: Creates timestamped backup before making changes"
        )
    
    return {
        'month_year': month_year,
        'include_closed': include_closed,
        'preview_only': preview_only,
        'create_backup': create_backup
    }

def extract_wip_data(wip_file):
    """Extract the exact fields we need from WIP Worksheet"""
    df = pd.read_excel(wip_file)
    
    # Get the columns by position (0-indexed)
    result = pd.DataFrame()
    result['Job Number'] = df.iloc[:, 0].astype(str).str.strip()  # Column A
    result['Job Name'] = df.iloc[:, 1]  # Column B
    result['Contract Amount'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)  # Column D
    result['Estimated Sub Labor'] = pd.to_numeric(df.iloc[:, 5], errors='coerce').fillna(0)  # Column F
    result['Estimated Material'] = pd.to_numeric(df.iloc[:, 6], errors='coerce').fillna(0)  # Column G
    
    return result

def extract_gl_data(gl_file):
    """Extract Labor Actual, Material Actual, and Amount Billed from GL"""
    df = pd.read_excel(gl_file)
    
    # Filter for relevant accounts
    account_filters = ['5040', '5030', '4020']
    mask = df['Account'].astype(str).str.contains('|'.join(account_filters), na=False)
    df = df[mask]
    
    # Clean job numbers
    df['Job Number'] = df['Job Number'].astype(str).str.strip()
    
    # Convert amounts
    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce').fillna(0)
    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce').fillna(0)
    df['Amount'] = df['Debit'] + df['Credit']
    
    # Calculate Amount Billed (K + L)
    amount_billed = 0
    if 'K' in df.columns and 'L' in df.columns:
        df['K'] = pd.to_numeric(df['K'], errors='coerce').fillna(0)
        df['L'] = pd.to_numeric(df['L'], errors='coerce').fillna(0)
        amount_billed = df['K'] + df['L']
    df['Amount Billed'] = amount_billed
    
    # Categorize by account type
    def get_account_type(account):
        if '5040' in str(account):
            return 'Labor'
        elif '5030' in str(account):
            return 'Material'
        return 'Other'
    
    df['Account Type'] = df['Account'].apply(get_account_type)
    
    # Aggregate by Job Number and Account Type
    labor_data = df[df['Account Type'] == 'Labor'].groupby('Job Number').agg({
        'Amount': 'sum',
        'Amount Billed': 'sum'
    }).reset_index()
    labor_data.rename(columns={'Amount': 'Labor Actual'}, inplace=True)
    
    material_data = df[df['Account Type'] == 'Material'].groupby('Job Number').agg({
        'Amount': 'sum'
    }).reset_index()
    material_data.rename(columns={'Amount': 'Material Actual'}, inplace=True)
    
    return labor_data, material_data

def create_final_data(wip_data, labor_data, material_data, include_closed=False):
    """Combine all data for final output"""
    
    # Filter out closed jobs if requested
    if not include_closed and 'Status' in wip_data.columns:
        wip_data = wip_data[wip_data['Status'] != 'Closed']
    
    # Merge with labor data
    final_5040 = wip_data.merge(labor_data, on='Job Number', how='left')
    final_5040['Labor Actual'] = final_5040['Labor Actual'].fillna(0)
    final_5040['Amount Billed'] = final_5040['Amount Billed'].fillna(0)
    
    # Merge with material data  
    final_5030 = wip_data.merge(material_data, on='Job Number', how='left')
    final_5030['Material Actual'] = final_5030['Material Actual'].fillna(0)
    
    return final_5040, final_5030

def update_excel_simple(master_file, data_5040, data_5030, month_year):
    """Update Excel with the exact data needed"""
    wb = load_workbook(master_file, keep_vba=True)
    
    # Get or create worksheet
    if month_year not in wb.sheetnames:
        ws = wb.create_sheet(month_year)
    else:
        ws = wb[month_year]
    
    # Write 5040 section
    ws['A1'] = '5040'
    row = 2
    for _, job in data_5040.iterrows():
        ws[f'A{row}'] = job['Job Number']
        ws[f'B{row}'] = job['Job Name']
        ws[f'C{row}'] = float(job['Contract Amount'])
        ws[f'D{row}'] = float(job['Estimated Sub Labor'])
        ws[f'E{row}'] = float(job['Labor Actual'])
        ws[f'F{row}'] = float(job['Amount Billed'])
        row += 1
    
    # Write 5030 section (leave some space)
    start_5030 = row + 5
    ws[f'A{start_5030}'] = '5030'
    row = start_5030 + 1
    for _, job in data_5030.iterrows():
        ws[f'A{row}'] = job['Job Number']
        ws[f'B{row}'] = job['Job Name']
        ws[f'C{row}'] = float(job['Estimated Material'])
        ws[f'D{row}'] = float(job['Material Actual'])
        row += 1
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def display_data_preview(merged_df, gl_aggregated):
    """Display preview of processed data"""
    st.markdown('<div class="section-header">👁️ Data Preview</div>', unsafe_allow_html=True)
    
    # Summary statistics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_jobs = len(merged_df)
        st.metric("Total Jobs", total_jobs)
    
    with col2:
        jobs_with_activity = len(merged_df[(merged_df['Sub Labor'] > 0) | (merged_df['Material'] > 0)])
        st.metric("Jobs with Activity", jobs_with_activity)
    
    with col3:
        large_variances = len(merged_df[
            (abs(merged_df.get('Sub Labor Variance', 0)) > 1000) |
            (abs(merged_df.get('Material Variance', 0)) > 1000)
        ])
        st.metric("Large Variances (>$1,000)", large_variances)
    
    with col4:
        closed_jobs = len(merged_df[merged_df['Status'] == 'Closed'])
        st.metric("Closed Jobs", closed_jobs)
    
    # Tabbed data display
    tab1, tab2, tab3 = st.tabs(["📊 Merged Data", "🔢 GL Aggregated", "⚠️ Large Variances"])
    
    with tab1:
        st.subheader("WIP Data Merged with GL Actuals")
        st.dataframe(
            merged_df,
            use_container_width=True,
            hide_index=True
        )
    
    with tab2:
        st.subheader("GL Data Aggregated by Job")
        st.dataframe(
            gl_aggregated,
            use_container_width=True,
            hide_index=True
        )
    
    with tab3:
        # Filter for large variances
        variance_mask = (
            (abs(merged_df.get('Sub Labor Variance', 0)) > 1000) |
            (abs(merged_df.get('Material Variance', 0)) > 1000)
        )
        large_variance_df = merged_df[variance_mask]
        
        if len(large_variance_df) > 0:
            st.subheader("Jobs with Variances > $1,000")
            st.dataframe(
                large_variance_df,
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No jobs with variances exceeding $1,000")

def display_excel_preview(options):
    """Display preview of Excel sections that would be updated"""
    if not options['preview_only']:
        return
        
    st.markdown('<div class="section-header">📋 Excel Preview</div>', unsafe_allow_html=True)
    
    try:
        # Load the master workbook
        master_file_bytes = st.session_state.files_uploaded['master_report'].getvalue()
        
        with st.spinner("Analyzing Master WIP Report structure..."):
            # Save to temporary file for openpyxl
            temp_path = "temp_master.xlsx"
            with open(temp_path, "wb") as f:
                f.write(master_file_bytes)
            
            # Load workbook and find sections
            wb = load_wip_workbook(temp_path)
            ws = find_or_create_monthly_tab(wb, options['month_year'])
            
            # Find sections
            section_markers = find_section_markers(ws, ["5040", "5030"])
            section_5040_row = section_markers.get("5040", (None, None))[0] if section_markers.get("5040") else None
            section_5030_row = section_markers.get("5030", (None, None))[0] if section_markers.get("5030") else None
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🔵 5040 Section (Sub Labor)")
                if section_5040_row:
                    st.success(f"✓ Found at row {section_5040_row}")
                    st.info("Section detected and ready for updates")
                else:
                    st.error("❌ Section not found")
            
            with col2:
                st.subheader("🟢 5030 Section (Material)")
                if section_5030_row:
                    st.success(f"✓ Found at row {section_5030_row}")
                    st.info("Section detected and ready for updates")
                else:
                    st.error("❌ Section not found")
            
            # Clean up temp file
            os.remove(temp_path)
            
    except Exception as e:
        st.error(f"Error analyzing Excel structure: {str(e)}")

def update_excel_file(merged_df, options):
    """Update the Excel file with processed data"""
    if options['preview_only']:
        st.info("Preview mode - no files will be updated")
        return None
        
    try:
        with st.spinner("Updating Master WIP Report..."):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Save master file temporarily (use absolute path to avoid directory issues)
            master_file_bytes = st.session_state.files_uploaded['master_report'].getvalue()
            temp_path = "/app/temp_master_update.xlsx"
            with open(temp_path, "wb") as f:
                f.write(master_file_bytes)
            
            # Backup will be handled by update_wip_report_v2 function
            
            # Load workbook
            status_text.text("📂 Loading workbook...")
            progress_bar.progress(40)
            wb = load_wip_workbook(temp_path)
            ws = find_or_create_monthly_tab(wb, options['month_year'])
            
            # Prepare data for update
            status_text.text("✍️ Preparing data for update...")
            progress_bar.progress(60)
            
            # Separate data for each section
            sub_labor_data = merged_df[merged_df['Sub Labor'] > 0][['Job Number', 'Sub Labor']].copy()
            
            material_data = merged_df[merged_df['Material'] > 0][['Job Number', 'Material']].copy()
            
            # Update both sections using the update function directly
            status_text.text("✍️ Updating Excel sections...")
            progress_bar.progress(80)
            
            # Close workbook before calling update function (to avoid conflicts)
            wb.close()
            
            # Call update function which handles everything internally
            result = update_wip_report_v2(temp_path, sub_labor_data, material_data, options['month_year'], options['create_backup'])
            
            # Check if update was successful
            if not result['success']:
                raise Exception(f"Update failed: {result.get('error', 'Unknown error')}")
            
            # Store backup info for display
            if result['backup_created']:
                st.session_state.backup_created = result['backup_created']
            
            # File has been updated by update_wip_report_v2
            status_text.text("💾 Finalizing updated report...")
            progress_bar.progress(90)
            
            # Read updated file for download using proper buffer handling
            import io
            with open(temp_path, "rb") as f:
                file_data = f.read()
            
            # Create BytesIO buffer for proper handling
            excel_buffer = io.BytesIO(file_data)
            excel_buffer.seek(0)  # Critical: reset to beginning
            updated_file_bytes = excel_buffer.getvalue()
            excel_buffer.close()
            
            progress_bar.progress(100)
            status_text.text("✅ Update complete!")
            
            # Clean up
            os.remove(temp_path)
            
            return updated_file_bytes
            
    except Exception as e:
        st.error(f"Error updating Excel file: {str(e)}")
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return None

def display_download_section(updated_file_bytes, merged_df):
    """Display download buttons for updated files"""
    st.markdown('<div class="section-header">📥 Download Results</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if updated_file_bytes:
            st.download_button(
                label="📊 Download Updated WIP Report",
                data=updated_file_bytes,
                file_name=f"WIP_Report_Updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download the updated Master WIP Report"
            )
        else:
            st.info("Updated report will appear here after processing")
    
    with col2:
        # Create validation report
        if merged_df is not None:
            validation_df = merged_df[
                (abs(merged_df.get('Sub Labor Variance', 0)) > 1000) |
                (abs(merged_df.get('Material Variance', 0)) > 1000)
            ].copy()
            
            if not validation_df.empty:
                # Convert to Excel bytes with proper buffer handling
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    validation_df.to_excel(writer, sheet_name='Large Variances', index=False)
                
                # Critical: seek to beginning before getting value
                output.seek(0)
                validation_bytes = output.getvalue()
                output.close()
                
                st.download_button(
                    label="⚠️ Download Validation Report",
                    data=validation_bytes,
                    file_name=f"WIP_Validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download report of jobs with large variances"
                )
            else:
                st.success("No large variances found - no validation report needed")

def main():
    """Main Streamlit application"""
    initialize_session_state()
    
    # Header
    st.markdown('<div class="main-header">WIP Report Automation Tool</div>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Sidebar with instructions
    with st.sidebar:
        st.header("📖 Instructions")
        st.markdown("""
        **Step 1:** Upload all three required Excel files
        
        **Step 2:** Configure processing options
        
        **Step 3:** Click Process to run automation
        
        **Step 4:** Review preview and download results
        
        ---
        
        **Files Required:**
        - **GL Inquiry**: Actual costs from Sage
        - **WIP Worksheet**: Job budgets from Sage  
        - **Master WIP Report**: File to be updated
        
        ---
        
        **Safety Features:**
        - ✅ Automatic backups
        - ✅ Formula preservation
        - ✅ Preview before update
        - ✅ Validation reports
        """)
    
    # Main content
    files_ready = display_file_upload_section()
    
    if files_ready:
        options = display_processing_options()
        
        # Process button
        if st.button("🚀 Process Data", type="primary", use_container_width=True):
            try:
                with st.spinner("Processing data..."):
                    # Load and process data
                    st.info("Loading GL Inquiry data...")
                    gl_df = load_and_process_gl_data(st.session_state.files_uploaded['gl_inquiry'])
                    st.success(f"Processed {len(gl_df)} GL records")
                    
                    st.info("Loading WIP Worksheet data...")
                    wip_df = load_wip_worksheet(st.session_state.files_uploaded['wip_worksheet'])
                    st.success(f"Loaded {len(wip_df)} WIP records")
                    
                    st.info("Merging data...")
                    merged_df = merge_data(wip_df, gl_df, include_closed=options['include_closed'])
                    st.success(f"Merged to {len(merged_df)} final records")
                    
                    # Display preview with ALL the new fields
                    st.subheader("Data Preview - All Fields")
                    preview_cols = ['Job Number', 'Job Description', 'Status', 
                                  'Contract Amount', 'Estimated Sub Labor', 'Sub Labor Actual',
                                  'Estimated Material', 'Material Actual', 'Amount Billed']
                    
                    display_df = merged_df[preview_cols].copy()
                    for col in ['Contract Amount', 'Estimated Sub Labor', 'Sub Labor Actual', 
                               'Estimated Material', 'Material Actual', 'Amount Billed']:
                        if col in display_df.columns:
                            display_df[col] = display_df[col].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "$0.00")
                    
                    st.dataframe(display_df)
                    
                    # Update Excel file
                    st.info("Updating WIP Report...")
                    updated_bytes = update_excel_simple(st.session_state.files_uploaded['master_report'], wip_df, gl_df, options['month_year'])
                    st.success("WIP Report updated successfully!")
                    
                    # Download button
                    st.download_button(
                        label="Download Updated WIP Report",
                        data=updated_bytes,
                        file_name=f"WIP_Report_{options['month_year'].replace(' ', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Summary
                    st.subheader("Processing Summary")
                    st.metric("Jobs Processed", len(merged_df))
                    st.metric("Total Contract Amount", f"${merged_df['Contract Amount'].sum():,.2f}")
                    st.metric("Total Sub Labor Actual", f"${merged_df['Sub Labor Actual'].sum():,.2f}")
                    st.metric("Total Material Actual", f"${merged_df['Material Actual'].sum():,.2f}")
                    st.metric("Total Amount Billed", f"${merged_df['Amount Billed'].sum():,.2f}")
                    
            except Exception as e:
                st.error(f"Error processing data: {str(e)}")
                logger.error(f"Processing error: {str(e)}")
    
    # Display results if processing is complete
    if st.session_state.processing_complete and st.session_state.processed_data:
        merged_df, gl_aggregated, options = st.session_state.processed_data
        
        # Data preview
        display_data_preview(merged_df, gl_aggregated)
        
        # Excel preview
        display_excel_preview(options)
        
        # Update Excel file if not preview only
        updated_file_bytes = None
        if not options['preview_only']:
            if st.button("✍️ Apply Updates to Excel", type="primary"):
                updated_file_bytes = update_excel_file(merged_df, options)
                if updated_file_bytes:
                    st.success("✅ Excel file updated successfully!")
                    if st.session_state.backup_created:
                        st.info(f"💾 Backup created: {st.session_state.backup_created}")
        
        # Download section
        display_download_section(updated_file_bytes, merged_df)
        
        # Reset button
        if st.button("🔄 Start Over"):
            for key in st.session_state.keys():
                del st.session_state[key]
            st.rerun()

if __name__ == "__main__":
    main() 