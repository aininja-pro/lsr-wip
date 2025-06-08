#!/usr/bin/env python3
"""
Ultra-Safe WIP Report Automation - Minimal Excel Updates
This version focuses on safe, data-only updates without complex preservation
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime
from pathlib import Path
import logging

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

def display_file_upload_section():
    """Display file upload widgets"""
    st.subheader("📁 File Upload")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.write("**Master WIP Report**")
        master_file = st.file_uploader(
            "Upload Master WIP Report",
            type=['xlsx', 'xlsm'],
            key='master_report'
        )
        if master_file:
            st.session_state.files_uploaded['master_report'] = master_file
            st.success(f"✅ {master_file.name}")
    
    with col2:
        st.write("**WIP Worksheet Export**")
        wip_file = st.file_uploader(
            "Upload WIP Worksheet",
            type=['xlsx'],
            key='wip_worksheet'
        )
        if wip_file:
            st.session_state.files_uploaded['wip_worksheet'] = wip_file
            st.success(f"✅ {wip_file.name}")
    
    with col3:
        st.write("**GL Inquiry Export**")
        gl_file = st.file_uploader(
            "Upload GL Inquiry",
            type=['xlsx'],
            key='gl_inquiry'
        )
        if gl_file:
            st.session_state.files_uploaded['gl_inquiry'] = gl_file
            st.success(f"✅ {gl_file.name}")
    
    return len(st.session_state.files_uploaded) == 3

def process_gl_data(gl_file_bytes):
    """Process GL inquiry data"""
    try:
        # Load GL data
        gl_df = pd.read_excel(io.BytesIO(gl_file_bytes))
        
        # Simple column mapping
        if 'JobNumber' in gl_df.columns:
            gl_df = gl_df.rename(columns={'JobNumber': 'Job Number'})
        
        # Filter for accounts containing our target strings
        account_filters = ['5040', '5030', '4020']
        mask = gl_df['Account'].astype(str).str.contains('|'.join(account_filters), na=False)
        filtered_gl = gl_df[mask].copy()
        
        # Compute Amount = Debit + Credit
        filtered_gl['Amount'] = filtered_gl['Debit'].fillna(0) + filtered_gl['Credit'].fillna(0)
        
        # Group by Job Number and sum amounts
        gl_summary = filtered_gl.groupby('Job Number')['Amount'].sum().reset_index()
        
        st.info(f"✅ Processed {len(gl_summary)} GL job entries from {len(filtered_gl)} records")
        return gl_summary
        
    except Exception as e:
        st.error(f"Error processing GL data: {str(e)}")
        return None

def process_wip_data(wip_file_bytes, gl_summary, include_closed=False):
    """Process WIP worksheet and merge with GL data"""
    try:
        # Load WIP data
        wip_df = pd.read_excel(io.BytesIO(wip_file_bytes))
        
        # Simple column mapping
        if 'JobNumber' in wip_df.columns:
            wip_df = wip_df.rename(columns={'JobNumber': 'Job Number'})
        
        # Trim job numbers
        wip_df['Job Number'] = wip_df['Job Number'].astype(str).str.strip()
        gl_summary['Job Number'] = gl_summary['Job Number'].astype(str).str.strip()
        
        # Filter closed jobs if requested
        if not include_closed and 'Status' in wip_df.columns:
            original_count = len(wip_df)
            wip_df = wip_df[wip_df['Status'].astype(str).str.upper() != 'CLOSED']
            st.info(f"📋 Filtered out {original_count - len(wip_df)} closed jobs")
        
        # Merge with GL data
        merged_df = wip_df.merge(gl_summary, on='Job Number', how='left', suffixes=('', '_GL'))
        merged_df['Amount'] = merged_df['Amount'].fillna(0)
        
        st.info(f"✅ Merged {len(wip_df)} WIP jobs with {len(gl_summary)} GL entries")
        return merged_df
        
    except Exception as e:
        st.error(f"Error processing WIP data: {str(e)}")
        return None

def create_simple_csv_output(merged_df):
    """Create a simple CSV output instead of Excel modification"""
    try:
        # Create output data focused on key fields
        output_df = merged_df[['Job Number', 'Amount']].copy()
        output_df['Labor_Cost'] = output_df['Amount'] * 0.6  # Example split
        output_df['Material_Cost'] = output_df['Amount'] * 0.4  # Example split
        
        # Convert to CSV
        csv_buffer = io.StringIO()
        output_df.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        csv_buffer.close()
        
        return csv_data
        
    except Exception as e:
        st.error(f"Error creating CSV output: {str(e)}")
        return None

def main():
    """Main application"""
    st.set_page_config(
        page_title="WIP Report Automation - Safe Mode",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 WIP Report Automation - Safe Mode")
    st.write("Ultra-safe data processing without Excel file modification")
    
    initialize_session_state()
    
    # File upload section
    files_ready = display_file_upload_section()
    
    if not files_ready:
        st.info("👆 Please upload all three required files to continue")
        return
    
    # Processing options
    st.subheader("⚙️ Processing Options")
    col1, col2 = st.columns(2)
    
    with col1:
        include_closed = st.checkbox("Include Closed Jobs", value=False)
    
    with col2:
        month_year = st.text_input("Month/Year", value="Apr 25")
    
    # Process button
    if st.button("🔄 Process Data", type="primary"):
        with st.spinner("Processing data..."):
            
            # Process GL data
            gl_file_bytes = st.session_state.files_uploaded['gl_inquiry'].getvalue()
            gl_summary = process_gl_data(gl_file_bytes)
            
            if gl_summary is None:
                st.error("❌ Failed to process GL data")
                return
            
            # Process WIP data
            wip_file_bytes = st.session_state.files_uploaded['wip_worksheet'].getvalue()
            merged_df = process_wip_data(wip_file_bytes, gl_summary, include_closed)
            
            if merged_df is None:
                st.error("❌ Failed to process WIP data")
                return
            
            # Store results
            st.session_state.merged_data = merged_df
            st.session_state.processing_complete = True
            
            st.success("✅ Data processing complete!")
    
    # Display results if processing is complete
    if st.session_state.processing_complete and st.session_state.merged_data is not None:
        st.subheader("📊 Processing Results")
        
        merged_df = st.session_state.merged_data
        
        # Show summary statistics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Jobs", len(merged_df))
        with col2:
            jobs_with_data = len(merged_df[merged_df['Amount'] > 0])
            st.metric("Jobs with GL Data", jobs_with_data)
        with col3:
            total_amount = merged_df['Amount'].sum()
            st.metric("Total Amount", f"${total_amount:,.2f}")
        
        # Show data preview
        st.subheader("🔍 Data Preview")
        st.dataframe(merged_df.head(10))
        
        # Download options
        st.subheader("📥 Download Results")
        
        # CSV Download
        csv_data = create_simple_csv_output(merged_df)
        if csv_data:
            st.download_button(
                "📄 Download CSV Results",
                data=csv_data,
                file_name=f"WIP_Results_{month_year.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        # Excel Download (Simple)
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            merged_df.to_excel(writer, sheet_name='Results', index=False)
        excel_buffer.seek(0)
        
        st.download_button(
            "📊 Download Excel Results",
            data=excel_buffer.getvalue(),
            file_name=f"WIP_Results_{month_year.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Instructions for manual update
        st.subheader("📋 Manual Update Instructions")
        st.info("""
        **Safe Manual Update Process:**
        1. Download the CSV or Excel results above
        2. Open your original Master WIP Report in Excel
        3. Create a backup copy first
        4. Manually copy the data from the results into the appropriate sections
        5. This ensures no corruption of your original file
        
        **This approach is 100% safe but requires manual copying of the data.**
        """)
        
        # Reset button
        if st.button("🔄 Start Over"):
            st.session_state.clear()
            st.experimental_rerun()

if __name__ == "__main__":
    main() 