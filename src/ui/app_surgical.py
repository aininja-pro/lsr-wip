"""
WIP Report Automation Tool - Surgical Excel Edition

This version uses a surgical ZIP/XML approach to modify Excel files,
avoiding the 56KB data loss issue with openpyxl.

Key Features:
- Preserves ALL Excel content: formulas, formatting, VBA, metadata
- Surgical cell-level updates only
- Zero data corruption
- Enterprise-grade file handling
"""

import streamlit as st
import pandas as pd
import io
import logging
from datetime import datetime

# Import our modules
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from data_processing.aggregation import (
    filter_gl_accounts, 
    compute_amounts, 
    aggregate_gl_data
)
from data_processing.merge_data import merge_wip_with_gl
from data_processing.excel_surgical import (
    update_wip_report_surgical, 
    create_backup_from_bytes,
    find_section_locations
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page config
st.set_page_config(
    page_title="WIP Report Automation - Surgical Edition",
    page_icon="🏗️",
    layout="wide"
)

def load_and_validate_file(uploaded_file, file_type="Excel"):
    """Load and validate uploaded file"""
    if uploaded_file is None:
        return None, None
        
    try:
        file_bytes = uploaded_file.read()
        
        if file_type == "Excel":
            # Try to load as Excel to validate
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
            return file_bytes, df
        else:
            return file_bytes, None
            
    except Exception as e:
        st.error(f"Error loading {file_type} file: {str(e)}")
        return None, None

def display_file_upload_section():
    """Display file upload widgets and return file data"""
    st.header("📁 File Upload")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Master WIP Report")
        master_file = st.file_uploader(
            "Upload Excel file (.xlsx or .xlsm)",
            type=['xlsx', 'xlsm'],
            key="master"
        )
        
        if master_file:
            st.success(f"✅ {master_file.name}")
            master_bytes, _ = load_and_validate_file(master_file, "Master")
        else:
            master_bytes = None
    
    with col2:
        st.subheader("WIP Worksheet Export")  
        wip_file = st.file_uploader(
            "Upload Excel file (.xlsx)",
            type=['xlsx'],
            key="wip"
        )
        
        if wip_file:
            st.success(f"✅ {wip_file.name}")
            wip_bytes, wip_df = load_and_validate_file(wip_file, "Excel")
        else:
            wip_bytes, wip_df = None, None
    
    with col3:
        st.subheader("GL Inquiry Export")
        gl_file = st.file_uploader(
            "Upload Excel file (.xlsx)",
            type=['xlsx'], 
            key="gl"
        )
        
        if gl_file:
            st.success(f"✅ {gl_file.name}")
            gl_bytes, gl_df = load_and_validate_file(gl_file, "Excel")
        else:
            gl_bytes, gl_df = None, None
    
    return master_bytes, wip_bytes, gl_bytes

def display_settings_section():
    """Display settings and options"""
    st.header("⚙️ Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Month/Year selector
        current_date = datetime.now()
        selected_date = st.date_input(
            "Select Month/Year for Report",
            value=current_date,
            help="Choose the month and year for the WIP report tab"
        )
        month_year = selected_date.strftime("%b %y")
        
    with col2:
        # Processing options
        include_closed = st.checkbox(
            "Include Closed Jobs",
            value=False,
            help="Include jobs marked as 'Closed' in processing"
        )
    
    return month_year, include_closed

def process_data(wip_bytes, gl_bytes, include_closed):
    """Process the data using our existing functions"""
    try:
        with st.spinner("Processing GL data..."):
            # Load GL inquiry from bytes
            gl_df = pd.read_excel(io.BytesIO(gl_bytes))
            
            # Apply column mapping (like load_gl_inquiry does)
            column_variations = {
                'Account': ['Account', 'Account Number', 'Acct', 'GL Account'],
                'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number'],
                'Debit': ['Debit', 'Debit Amount', 'DR', 'Dr'],
                'Credit': ['Credit', 'Credit Amount', 'CR', 'Cr']
            }
            
            # Map column names to standard names
            column_mapping = {}
            for standard_name, variations in column_variations.items():
                found_column = None
                for variation in variations:
                    if variation in gl_df.columns:
                        found_column = variation
                        break
                
                if found_column:
                    column_mapping[found_column] = standard_name
                else:
                    raise ValueError(f"Required column '{standard_name}' not found. Available columns: {list(gl_df.columns)}")
            
            # Rename columns to standard names
            gl_df = gl_df.rename(columns=column_mapping)
            
            # Process GL data step by step (instead of using the file-path version)
            filtered_gl = filter_gl_accounts(gl_df)
            amounts_gl = compute_amounts(filtered_gl)
            gl_summary = aggregate_gl_data(amounts_gl)
            
            st.info(f"✅ Processed {len(gl_summary)} GL entries")
            
        with st.spinner("Merging data..."):
            # Load WIP worksheet from bytes
            wip_df = pd.read_excel(io.BytesIO(wip_bytes))
            
            # Apply column mapping for WIP worksheet
            wip_column_variations = {
                'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number', 'Project No'],
                'Status': ['Status', 'Job Status', 'Project Status', 'State'],
                'Job Name': ['Job Name', 'Project Name', 'Description', 'Job Description'],
                'Budget Material': ['Budget Material', 'Material Budget', 'Mat Budget', 'Budget Mat'],
                'Budget Labor': ['Budget Labor', 'Labor Budget', 'Lab Budget', 'Budget Lab'],
                'Actual Material': ['Actual Material', 'Material Actual', 'Mat Actual', 'Actual Mat'],
                'Actual Labor': ['Actual Labor', 'Labor Actual', 'Lab Actual', 'Actual Lab']
            }
            
            # Map WIP column names to standard names
            wip_column_mapping = {}
            for standard_name, variations in wip_column_variations.items():
                found_column = None
                for variation in variations:
                    if variation in wip_df.columns:
                        found_column = variation
                        break
                
                if found_column:
                    wip_column_mapping[found_column] = standard_name
                else:
                    # Some columns might be optional, only require Job Number and Status
                    if standard_name in ['Job Number', 'Status']:
                        raise ValueError(f"Required WIP column '{standard_name}' not found. Available columns: {list(wip_df.columns)}")
            
            # Rename WIP columns to standard names
            wip_df = wip_df.rename(columns=wip_column_mapping)
            
            merged_df = merge_wip_with_gl(wip_df, gl_summary, include_closed)
            st.info(f"✅ Merged data for {len(merged_df)} jobs")
            
        return merged_df
        
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        logger.error(f"Processing error: {e}")
        return None

def display_preview_section(merged_df):
    """Display data preview"""
    if merged_df is not None and not merged_df.empty:
        st.header("👀 Data Preview")
        
        # Summary stats
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Jobs", len(merged_df))
        with col2:
            total_labor = merged_df.get('Sub Labor', pd.Series([0])).sum()
            st.metric("Total Sub Labor", f"${total_labor:,.2f}")
        with col3:
            total_material = merged_df.get('Material', pd.Series([0])).sum()
            st.metric("Total Material", f"${total_material:,.2f}")
        with col4:
            # Count non-zero entries
            non_zero_labor = (merged_df.get('Sub Labor', pd.Series([0])) != 0).sum()
            non_zero_material = (merged_df.get('Material', pd.Series([0])) != 0).sum()
            st.metric("Active Entries", f"{non_zero_labor + non_zero_material}")
        
        # Data table
        st.subheader("Merged Data")
        st.dataframe(merged_df, use_container_width=True)
        
        return True
    return False

def display_processing_section(master_bytes, merged_df, month_year):
    """Display processing section and handle Excel updates"""
    if master_bytes and merged_df is not None:
        st.header("🔄 Excel Processing")
        
        # Show section detection first
        with st.spinner("Detecting Excel sections..."):
            row_5040, row_5030 = find_section_locations(master_bytes, month_year)
            
            if row_5040:
                st.success(f"✅ Found 5040 section at row {row_5040}")
            else:
                st.warning("⚠️ Could not find 5040 section")
                
            if row_5030:
                st.success(f"✅ Found 5030 section at row {row_5030}")
            else:
                st.warning("⚠️ Could not find 5030 section")
        
        if not row_5040 and not row_5030:
            st.error("❌ No sections found. Cannot proceed with update.")
            return None
        
        # Process button
        if st.button("🚀 Process WIP Report", type="primary"):
            
            # Create backup first
            with st.spinner("Creating backup..."):
                backup_filename = create_backup_from_bytes(master_bytes)
                if backup_filename:
                    st.success(f"✅ Backup created: {backup_filename}")
                else:
                    st.error("❌ Failed to create backup")
                    return None
            
            # Perform surgical update
            with st.spinner("Performing surgical Excel update..."):
                try:
                    updated_bytes = update_wip_report_surgical(
                        master_bytes, 
                        merged_df, 
                        month_year
                    )
                    
                    # File size comparison
                    original_size = len(master_bytes)
                    updated_size = len(updated_bytes)
                    size_diff = updated_size - original_size
                    
                    st.success("✅ Surgical update completed!")
                    
                    # Show size info
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.info(f"📁 Original: {original_size:,} bytes")
                    with col2:
                        st.info(f"📁 Updated: {updated_size:,} bytes")
                    with col3:
                        if abs(size_diff) < 100:
                            st.success(f"📁 Difference: {size_diff:+,} bytes ✅")
                        else:
                            st.info(f"📁 Difference: {size_diff:+,} bytes")
                    
                    return updated_bytes
                    
                except Exception as e:
                    st.error(f"❌ Surgical update failed: {str(e)}")
                    logger.error(f"Surgical update error: {e}")
                    return None
    
    return None

def display_download_section(updated_bytes, merged_df):
    """Display download section"""
    if updated_bytes is not None:
        st.header("📥 Download Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📊 Download Updated WIP Report",
                data=updated_bytes,
                file_name="WIP_Report_Updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        
        with col2:
            # Create simple validation report  
            if merged_df is not None:
                # Create validation DataFrame
                validation_data = []
                for _, row in merged_df.iterrows():
                    job = row.get('Job Number', '')
                    material = row.get('Material', 0) or 0
                    labor = row.get('Sub Labor', 0) or 0
                    
                    if material > 1000 or labor > 1000:
                        validation_data.append({
                            'Job Number': job,
                            'Material': material,
                            'Sub Labor': labor,
                            'Flag': 'High Value' if (material > 1000 or labor > 1000) else 'Normal'
                        })
                
                if validation_data:
                    validation_df = pd.DataFrame(validation_data)
                    validation_buffer = io.BytesIO()
                    with pd.ExcelWriter(validation_buffer, engine='xlsxwriter') as writer:
                        validation_df.to_excel(writer, index=False, sheet_name='Validation')
                    validation_buffer.seek(0)
                    
                    st.download_button(
                        label="📋 Download Validation Report",
                        data=validation_buffer.getvalue(),
                        file_name="WIP_Validation_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

def main():
    """Main application"""
    # Title
    st.title("🏗️ WIP Report Automation Tool")
    st.markdown("### Surgical Excel Edition - Zero Data Loss")
    
    # File uploads
    master_bytes, wip_bytes, gl_bytes = display_file_upload_section()
    
    # Check if we have the required files
    if not all([master_bytes, wip_bytes, gl_bytes]):
        st.info("👆 Please upload all three Excel files to continue")
        return
    
    # Settings
    month_year, include_closed = display_settings_section()
    
    # Process data
    merged_df = process_data(wip_bytes, gl_bytes, include_closed)
    
    # Preview
    if display_preview_section(merged_df):
        # Processing
        updated_bytes = display_processing_section(master_bytes, merged_df, month_year)
        
        # Downloads
        display_download_section(updated_bytes, merged_df)

if __name__ == "__main__":
    main() 