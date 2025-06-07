#!/usr/bin/env python3
"""
Test script to validate section positions and job counts in the April 25 tab
"""

import pandas as pd
from pathlib import Path
import sys
import os

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from data_processing.excel_integration_v2 import (
    load_wip_workbook, 
    find_section_markers
)

def validate_april_25_tab():
    """Validate the April 25 tab structure and job counts."""
    print("🔍 Validating April 25 Tab Structure")
    print("=" * 50)
    
    # Load the Master WIP Report
    master_report_file = Path("test_data") / "Master WIP Report.xlsx"
    
    if not master_report_file.exists():
        print(f"❌ Master WIP Report not found: {master_report_file}")
        return False
    
    try:
        # Load workbook
        workbook = load_wip_workbook(str(master_report_file))
        print(f"✅ Loaded workbook with sheets: {workbook.sheetnames}")
        
        # Check if April 25 tab exists
        if 'April 25' not in workbook.sheetnames:
            print("❌ 'April 25' tab not found")
            return False
        
        # Get April 25 worksheet
        april_ws = workbook['April 25']
        print(f"✅ Found 'April 25' worksheet")
        
        # Find section markers in April 25
        print("\n🔍 Searching for section markers in April 25...")
        markers = find_section_markers(april_ws, ['5040', '5030'])
        print(f"Section markers found: {markers}")
        
        # Manual verification - let's scan the worksheet more thoroughly
        print("\n📊 Manual verification - scanning worksheet...")
        
        # Look for 5040 section
        print("Looking for 5040 section (Sub Labor)...")
        for row in range(1, 20):  # Check first 20 rows
            for col in range(1, 10):  # Check first 10 columns
                cell = april_ws.cell(row=row, column=col)
                if cell.value and '5040' in str(cell.value):
                    print(f"   Found '5040' at row {row}, col {col}: '{cell.value}'")
        
        # Look for 5030 section
        print("Looking for 5030 section (Material)...")
        for row in range(50, 80):  # Check around row 69
            for col in range(1, 10):  # Check first 10 columns
                cell = april_ws.cell(row=row, column=col)
                if cell.value and '5030' in str(cell.value):
                    print(f"   Found '5030' at row {row}, col {col}: '{cell.value}'")
        
        # Count jobs in 5040 section (starting around row 3)
        print(f"\n📋 Counting jobs in 5040 section (starting around row 3)...")
        job_count_5040 = 0
        for row in range(3, 30):  # Check rows 3-30
            job_cell = april_ws.cell(row=row, column=1)  # Column A
            if job_cell.value and str(job_cell.value).strip():
                # Check if it looks like a job number
                job_value = str(job_cell.value).strip()
                if job_value and not job_value.lower() in ['job#', 'job', 'total', '']:
                    print(f"   Row {row}, Col A: '{job_value}'")
                    job_count_5040 += 1
            else:
                # Stop counting when we hit empty rows
                consecutive_empty = 0
                for check_row in range(row, row + 3):
                    check_cell = april_ws.cell(row=check_row, column=1)
                    if not check_cell.value or not str(check_cell.value).strip():
                        consecutive_empty += 1
                if consecutive_empty >= 3:
                    break
        
        print(f"   Total jobs in 5040 section: {job_count_5040}")
        
        # Count jobs in 5030 section (starting around row 70)
        print(f"\n📋 Counting jobs in 5030 section (starting around row 70)...")
        job_count_5030 = 0
        for row in range(70, 120):  # Check rows 70-120
            desc_cell = april_ws.cell(row=row, column=1)  # Column A for descriptions
            if desc_cell.value and str(desc_cell.value).strip():
                # Check if it looks like a job description
                desc_value = str(desc_cell.value).strip()
                if desc_value and not desc_value.lower() in ['job description', 'total', '']:
                    print(f"   Row {row}, Col A: '{desc_value}'")
                    job_count_5030 += 1
            else:
                # Stop counting when we hit empty rows
                consecutive_empty = 0
                for check_row in range(row, row + 3):
                    check_cell = april_ws.cell(row=check_row, column=1)
                    if not check_cell.value or not str(check_cell.value).strip():
                        consecutive_empty += 1
                if consecutive_empty >= 3:
                    break
        
        print(f"   Total jobs in 5030 section: {job_count_5030}")
        
        print(f"\n📋 SUMMARY:")
        print(f"   • 5040 section jobs: {job_count_5040}")
        print(f"   • 5030 section jobs: {job_count_5030}")
        print(f"   • User reported: 5 jobs in 5040, 10 jobs in 5030")
        
        if job_count_5040 == 5 and job_count_5030 == 10:
            print("✅ PERFECT MATCH with user's count!")
        else:
            print("⚠️  Count mismatch - need to adjust our detection logic")
        
        return True
        
    except Exception as e:
        print(f"❌ ERROR during validation: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    validate_april_25_tab() 