#!/usr/bin/env python3
"""
Test script to validate our WIP automation modules with real client data.
This will help us identify any issues before building the Streamlit interface.
"""

import pandas as pd
from pathlib import Path
import sys
import os

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from data_processing.aggregation import process_gl_inquiry
from data_processing.merge_data import process_wip_merge
from data_processing.excel_integration import (
    load_wip_workbook, 
    find_or_create_monthly_tab,
    find_section_markers,
    get_existing_data_from_section
)
from data_processing.column_mapping import map_columns_for_file_type

def test_real_data():
    """Test our modules with real client data."""
    print("🚀 Testing WIP Automation with Real Client Data")
    print("=" * 60)
    
    # File paths
    test_data_dir = Path("test_data")
    gl_file = test_data_dir / "GL Inquiry Export.xlsx"
    wip_worksheet_file = test_data_dir / "WIP Worksheet Export.xlsx"
    master_report_file = test_data_dir / "Master WIP Report.xlsx"
    
    # Check files exist
    for file_path, name in [(gl_file, "GL Inquiry"), (wip_worksheet_file, "WIP Worksheet"), (master_report_file, "Master WIP Report")]:
        if not file_path.exists():
            print(f"❌ {name} file not found: {file_path}")
            return False
        else:
            print(f"✅ Found {name}: {file_path}")
    
    print()
    
    try:
        # Test 1: Load and examine GL Inquiry Export
        print("📊 Step 1: Testing GL Inquiry processing...")
        gl_df = pd.read_excel(gl_file)
        print(f"   Raw GL data shape: {gl_df.shape}")
        print(f"   GL columns: {list(gl_df.columns)}")
        
        # Test column mapping
        gl_mapping = map_columns_for_file_type(list(gl_df.columns), 'gl_inquiry')
        print(f"   GL column mapping: {gl_mapping}")
        
        # Process GL data
        gl_aggregated = process_gl_inquiry(str(gl_file))
        print(f"   Aggregated GL shape: {gl_aggregated.shape}")
        print(f"   GL columns after processing: {list(gl_aggregated.columns)}")
        if 'Account Type' in gl_aggregated.columns:
            print(f"   GL accounts found: {gl_aggregated['Account Type'].unique()}")
        print(f"   GL jobs count: {gl_aggregated['Job Number'].nunique()}")
        print("   Sample GL aggregated data:")
        print(gl_aggregated.head(3))
        print()
        
        # Test 2: Load and examine WIP Worksheet
        print("📋 Step 2: Testing WIP Worksheet processing...")
        wip_df = pd.read_excel(wip_worksheet_file)
        print(f"   Raw WIP data shape: {wip_df.shape}")
        print(f"   WIP columns: {list(wip_df.columns)}")
        
        # Test column mapping
        wip_mapping = map_columns_for_file_type(list(wip_df.columns), 'wip_worksheet')
        print(f"   WIP column mapping: {wip_mapping}")
        print()
        
        # Test 3: Merge WIP with GL data
        print("🔗 Step 3: Testing WIP merge process...")
        merged_data = process_wip_merge(str(wip_worksheet_file), gl_aggregated)
        print(f"   Merged data shape: {merged_data.shape}")
        print(f"   Jobs in merged data: {merged_data['Job Number'].nunique()}")
        
        # Show sample of merged data
        print("   Sample merged data:")
        print(merged_data.head(3))
        print()
        
        # Test 4: Load Master WIP Report
        print("📈 Step 4: Testing Master WIP Report reading...")
        workbook = load_wip_workbook(str(master_report_file))
        print(f"   Workbook loaded successfully")
        print(f"   Sheet names: {workbook.sheetnames}")
        
        # Test finding monthly tab (try current month)
        from datetime import datetime
        current_month = datetime.now().strftime("%b %y")
        print(f"   Looking for monthly tab: {current_month}")
        
        monthly_ws = find_or_create_monthly_tab(workbook, current_month)
        print(f"   Monthly worksheet: {monthly_ws.title}")
        
        # Test finding section markers
        markers = find_section_markers(monthly_ws, ['5040', '5030'])
        print(f"   Section markers found: {markers}")
        
        # Test extracting existing data if sections exist
        if '5040' in markers:
            existing_5040 = get_existing_data_from_section(monthly_ws, '5040')
            print(f"   Existing 5040 data shape: {existing_5040.shape}")
            print("   Sample existing 5040 data:")
            print(existing_5040.head(3))
        
        if '5030' in markers:
            existing_5030 = get_existing_data_from_section(monthly_ws, '5030')
            print(f"   Existing 5030 data shape: {existing_5030.shape}")
            print("   Sample existing 5030 data:")
            print(existing_5030.head(3))
        
        print()
        print("🎉 SUCCESS: All real data tests passed!")
        print("✨ The WIP automation tool is ready for your data structure!")
        
        return True
        
    except Exception as e:
        print(f"❌ ERROR during testing: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_real_data()
    if success:
        print("\n🚀 Ready to proceed to Phase 4: Streamlit Interface!")
    else:
        print("\n⚠️  Issues found - need to fix before proceeding") 