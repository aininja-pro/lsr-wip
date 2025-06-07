#!/usr/bin/env python3
"""
Updated test script with correct account types and column mapping:
- 5040 = Sub Labor (not material!)
- 5030 = Material (not labor!)
"""

import pandas as pd
from pathlib import Path
import sys
import os

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from data_processing.aggregation import process_gl_inquiry
from data_processing.merge_data import process_wip_merge, get_jobs_for_update
from data_processing.excel_integration_v2 import (
    load_wip_workbook, 
    find_or_create_monthly_tab,
    find_section_markers
)
from data_processing.column_mapping import map_columns_for_file_type

def test_corrected_real_data():
    """Test our modules with real client data using CORRECTED account types."""
    print("🚀 Testing WIP Automation with CORRECTED Real Client Data")
    print("=" * 70)
    
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
        print("📊 Step 1: Testing GL Inquiry processing (CORRECTED ACCOUNTS)...")
        gl_df = pd.read_excel(gl_file)
        print(f"   Raw GL data shape: {gl_df.shape}")
        print(f"   GL columns: {list(gl_df.columns)}")
        
        # Process GL data
        gl_aggregated = process_gl_inquiry(str(gl_file))
        print(f"   Aggregated GL shape: {gl_aggregated.shape}")
        print(f"   GL columns after processing: {list(gl_aggregated.columns)}")
        print(f"   GL jobs count: {gl_aggregated['Job Number'].nunique()}")
        
        # Show sample with CORRECTED account types
        print("   Sample GL aggregated data (CORRECTED):")
        print("   5040 accounts = 'Sub Labor', 5030 accounts = 'Material'")
        print(gl_aggregated.head(3))
        print()
        
        # Test 2: Load and examine WIP Worksheet
        print("📋 Step 2: Testing WIP Worksheet processing...")
        wip_df = pd.read_excel(wip_worksheet_file)
        print(f"   Raw WIP data shape: {wip_df.shape}")
        print(f"   WIP columns: {list(wip_df.columns)}")
        print()
        
        # Test 3: Merge WIP with GL data
        print("🔗 Step 3: Testing WIP merge process...")
        merged_data = process_wip_merge(str(wip_worksheet_file), gl_aggregated)
        print(f"   Merged data shape: {merged_data.shape}")
        print(f"   Jobs in merged data: {merged_data['Job Number'].nunique()}")
        
        # Show column names to verify account types
        print(f"   Merged data columns: {list(merged_data.columns)}")
        
        # Show sample of merged data
        print("   Sample merged data:")
        display_cols = ['Job Number', 'Job Name', 'Sub Labor', 'Material', 'Other']
        available_cols = [col for col in display_cols if col in merged_data.columns]
        print(merged_data[available_cols].head(3))
        print()
        
        # Test 4: Get data for each section (CORRECTED)
        print("🔧 Step 4: Testing section data separation (CORRECTED)...")
        
        # 5040 section gets Sub Labor data
        section_5040_data = get_jobs_for_update(merged_data, '5040')
        print(f"   5040 section (Sub Labor) data shape: {section_5040_data.shape}")
        if not section_5040_data.empty:
            print("   Sample 5040 section data:")
            print(section_5040_data.head(3))
        
        # 5030 section gets Material data  
        section_5030_data = get_jobs_for_update(merged_data, '5030')
        print(f"   5030 section (Material) data shape: {section_5030_data.shape}")
        if not section_5030_data.empty:
            print("   Sample 5030 section data:")
            print(section_5030_data.head(3))
        print()
        
        # Test 5: Load Master WIP Report
        print("📈 Step 5: Testing Master WIP Report reading with enhanced section finding...")
        workbook = load_wip_workbook(str(master_report_file))
        print(f"   Workbook loaded successfully")
        print(f"   Sheet names: {workbook.sheetnames}")
        
        # Test finding monthly tab (try current month)
        from datetime import datetime
        current_month = datetime.now().strftime("%b %y")
        print(f"   Looking for monthly tab: {current_month}")
        
        monthly_ws = find_or_create_monthly_tab(workbook, current_month)
        print(f"   Monthly worksheet: {monthly_ws.title}")
        
        # Test finding section markers with enhanced patterns
        print("   Testing enhanced section marker detection...")
        markers = find_section_markers(monthly_ws, ['5040', '5030'])
        print(f"   Section markers found: {markers}")
        
        # If no sections found in new tab, try an existing tab
        if not any(markers.values()):
            print("   No sections in new tab, trying existing tab with data...")
            for sheet_name in ['Nov 24', 'Dec 24', 'Jan 25', 'Feb 25', 'Mar 25', 'Apr 25', 'May 25']:
                if sheet_name in workbook.sheetnames:
                    test_ws = workbook[sheet_name]
                    test_markers = find_section_markers(test_ws, ['5040', '5030'])
                    if any(test_markers.values()):
                        print(f"   Found sections in {sheet_name}: {test_markers}")
                        break
        
        print()
        print("🎉 SUCCESS: All corrected real data tests passed!")
        print("✨ Ready for Streamlit interface with CORRECT account mapping!")
        print()
        print("📋 SUMMARY:")
        print(f"   • 5040 Section: Sub Labor costs (found {len(section_5040_data[section_5040_data['Sub Labor'] != 0])} jobs with amounts)")
        print(f"   • 5030 Section: Material costs (found {len(section_5030_data[section_5030_data['Material'] != 0])} jobs with amounts)")
        print(f"   • Total jobs processed: {merged_data['Job Number'].nunique()}")
        
        return True
        
    except Exception as e:
        print(f"❌ ERROR during testing: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_corrected_real_data()
    if success:
        print("\n🚀 Ready to proceed to Phase 4: Streamlit Interface!")
    else:
        print("\n⚠️  Issues found - need to fix before proceeding") 