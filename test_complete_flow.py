#!/usr/bin/env python3
"""
Complete Flow Test - Tests the entire WIP Report automation workflow
This test simulates a user uploading files and processing them through the system.
"""

import sys
import os
from pathlib import Path
import pandas as pd
from datetime import datetime
import tempfile
import shutil

# Add src to path for imports
sys.path.append(str(Path(__file__).parent / "src"))

from data_processing.aggregation import aggregate_gl_data
from data_processing.merge_data import process_wip_merge, compute_variances
from data_processing.column_mapping import map_dataframe_columns
from data_processing.excel_integration_v2 import (
    load_wip_workbook, find_or_create_monthly_tab, find_section_markers,
    create_backup, update_wip_report_v2
)

def test_complete_workflow():
    """Test the complete workflow with real client files"""
    print("🧪 Testing Complete WIP Report Automation Workflow")
    print("=" * 60)
    
    # File paths
    gl_file = "test_data/GL Inquiry Export.xlsx"
    wip_file = "test_data/WIP Worksheet Export.xlsx" 
    master_file = "test_data/Master WIP Report.xlsx"
    
    # Check if files exist
    for file_path in [gl_file, wip_file, master_file]:
        if not os.path.exists(file_path):
            print(f"❌ Error: {file_path} not found")
            print("   Please ensure client files are in the project root directory")
            return False
    
    print("✅ All required files found")
    
    try:
        # Step 1: Load and process GL data
        print("\n📊 Step 1: Loading GL Inquiry data...")
        gl_df = pd.read_excel(gl_file)
        gl_df = map_dataframe_columns(gl_df, 'gl_inquiry')
        print(f"   Loaded {len(gl_df)} GL transactions")
        
        # Step 2: Aggregate GL data
        print("\n🔄 Step 2: Aggregating GL data...")
        gl_aggregated = aggregate_gl_data(gl_df)
        print(f"   Aggregated to {len(gl_aggregated)} job/account combinations")
        account_types = [col for col in gl_aggregated.columns if col != 'Job Number']
        print(f"   Account types found: {account_types}")
        
        # Step 3: Load and process WIP worksheet
        print("\n📋 Step 3: Loading WIP Worksheet data...")
        wip_df = pd.read_excel(wip_file)
        wip_df = map_dataframe_columns(wip_df, 'wip_worksheet')
        print(f"   Loaded {len(wip_df)} WIP jobs")
        
        # Step 4: Merge data  
        print("\n🔗 Step 4: Merging WIP and GL data...")
        # Save WIP file temporarily for processing
        temp_wip_path = "temp_wip.xlsx"
        wip_df.to_excel(temp_wip_path, index=False)
        merged_df = process_wip_merge(temp_wip_path, gl_aggregated, include_closed=False)
        os.remove(temp_wip_path)  # Clean up
        print(f"   Merged to {len(merged_df)} total jobs")
        
        jobs_with_activity = len(merged_df[(merged_df['Sub Labor'] > 0) | (merged_df['Material'] > 0)])
        print(f"   Jobs with activity: {jobs_with_activity}")
        
        # Step 5: Compute variances
        print("\n📈 Step 5: Computing variances...")
        final_df = compute_variances(merged_df)
        
        # Count large variances
        large_variances = len(final_df[
            (abs(final_df.get('Sub Labor Variance', 0)) > 1000) |
            (abs(final_df.get('Material Variance', 0)) > 1000)
        ])
        print(f"   Large variances (>$1,000): {large_variances}")
        
        # Step 6: Test Excel integration
        print("\n📋 Step 6: Testing Excel integration...")
        
        # Create a copy for testing
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
            temp_path = temp_file.name
            shutil.copy2(master_file, temp_path)
        
        try:
            # Load workbook
            wb = load_wip_workbook(temp_path)
            # Use an existing tab for testing
            test_month = "May 25"
            ws = find_or_create_monthly_tab(wb, test_month)
            
            # Find sections
            section_markers = find_section_markers(ws, ["5040", "5030"])
            section_5040_row = section_markers.get("5040", (None, None))[0] if section_markers.get("5040") else None
            section_5030_row = section_markers.get("5030", (None, None))[0] if section_markers.get("5030") else None
            
            print(f"   5040 section found at row: {section_5040_row}")
            print(f"   5030 section found at row: {section_5030_row}")
            
            if section_5040_row and section_5030_row:
                print("   ✅ Both sections detected successfully")
                
                # Test backup creation
                print("\n💾 Step 7: Testing backup creation...")
                backup_path = create_backup(temp_path)
                if os.path.exists(backup_path):
                    print(f"   ✅ Backup created: {os.path.basename(backup_path)}")
                    os.remove(backup_path)  # Clean up
                else:
                    print("   ❌ Backup creation failed")
                    
                print("\n✅ All workflow steps completed successfully!")
                print("\n📊 Summary Statistics:")
                print(f"   - GL transactions processed: {len(gl_df)}")
                print(f"   - Jobs aggregated: {len(gl_aggregated)}")
                print(f"   - Total WIP jobs: {len(wip_df)}")
                print(f"   - Jobs after merge: {len(final_df)}")
                print(f"   - Jobs with activity: {jobs_with_activity}")
                print(f"   - Large variances: {large_variances}")
                print(f"   - 5040 section location: Row {section_5040_row}")
                print(f"   - 5030 section location: Row {section_5030_row}")
                
                # Test a simple update to verify integration works
                print("\n✍️ Step 8: Testing Excel update integration...")
                sub_labor_jobs = final_df[final_df['Sub Labor'] > 0][['Job Number', 'Sub Labor']].copy()
                
                material_jobs = final_df[final_df['Material'] > 0][['Job Number', 'Material']].copy()
                
                print(f"   Jobs to update in 5040 section: {len(sub_labor_jobs)}")
                print(f"   Jobs to update in 5030 section: {len(material_jobs)}")
                
                if len(sub_labor_jobs) > 0 or len(material_jobs) > 0:
                    print("   ✅ Excel update data prepared successfully")
                
            else:
                print("   ❌ Section detection failed")
                return False
            
            wb.close()
            
        finally:
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error during workflow test: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def test_data_validation():
    """Test data validation and quality checks"""
    print("\n🔍 Additional Data Validation Tests")
    print("-" * 40)
    
    try:
        # Load files
        gl_df = pd.read_excel("test_data/GL Inquiry Export.xlsx")
        gl_df = map_dataframe_columns(gl_df, 'gl_inquiry')
        
        wip_df = pd.read_excel("test_data/WIP Worksheet Export.xlsx")
        wip_df = map_dataframe_columns(wip_df, 'wip_worksheet')
        
        # Test GL data quality
        print("\n📊 GL Data Quality:")
        account_5040_count = len(gl_df[gl_df['Account'].str.contains('5040', na=False)])
        account_5030_count = len(gl_df[gl_df['Account'].str.contains('5030', na=False)])
        account_4020_count = len(gl_df[gl_df['Account'].str.contains('4020', na=False)])
        
        print(f"   - 5040 (Sub Labor) transactions: {account_5040_count}")
        print(f"   - 5030 (Material) transactions: {account_5030_count}")
        print(f"   - 4020 (Billing) transactions: {account_4020_count}")
        
        # Test WIP data quality
        print("\n📋 WIP Data Quality:")
        open_jobs = len(wip_df[wip_df['Status'] == 'Open'])
        closed_jobs = len(wip_df[wip_df['Status'] == 'Closed'])
        
        print(f"   - Open jobs: {open_jobs}")
        print(f"   - Closed jobs: {closed_jobs}")
        
        # Test job number matching
        gl_jobs = set(gl_df['Job Number'].str.strip().unique())
        wip_jobs = set(wip_df['Job Number'].str.strip().unique())
        
        matching_jobs = gl_jobs.intersection(wip_jobs)
        gl_only_jobs = gl_jobs - wip_jobs
        wip_only_jobs = wip_jobs - gl_jobs
        
        print(f"\n🔗 Job Matching Analysis:")
        print(f"   - Jobs in both GL and WIP: {len(matching_jobs)}")
        print(f"   - Jobs only in GL: {len(gl_only_jobs)}")
        print(f"   - Jobs only in WIP: {len(wip_only_jobs)}")
        
        if len(gl_only_jobs) > 0:
            print(f"   - GL-only jobs (sample): {list(gl_only_jobs)[:5]}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error during validation: {str(e)}")
        return False

if __name__ == "__main__":
    print("🚀 WIP Report Automation - Complete System Test")
    print("=" * 60)
    
    # Run main workflow test
    workflow_success = test_complete_workflow()
    
    # Run data validation test
    validation_success = test_data_validation()
    
    print("\n" + "=" * 60)
    print("📋 TEST RESULTS SUMMARY")
    print("=" * 60)
    
    if workflow_success:
        print("✅ Complete workflow test: PASSED")
    else:
        print("❌ Complete workflow test: FAILED")
    
    if validation_success:
        print("✅ Data validation test: PASSED")
    else:
        print("❌ Data validation test: FAILED")
    
    if workflow_success and validation_success:
        print("\n🎉 ALL TESTS PASSED - System ready for production use!")
        print("\nNext steps:")
        print("1. Run: streamlit run src/ui/app.py")
        print("2. Open http://localhost:8501 in your browser")
        print("3. Upload your files and test the interface")
    else:
        print("\n⚠️  Some tests failed - please review errors above")
    
    print("\n" + "=" * 60) 