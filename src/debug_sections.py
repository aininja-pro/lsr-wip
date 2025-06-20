#!/usr/bin/env python3
"""
Debug script to analyze section detection issues
"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent))

from data_processing.excel_integration_v2 import (
    load_wip_workbook, 
    find_or_create_monthly_tab,
    find_section_markers
)
import logging

# Set up detailed logging
logging.basicConfig(level=logging.DEBUG, format='%(levelname)s:%(name)s:%(message)s')
logger = logging.getLogger(__name__)

# Add import and remove debug hardcode
import pandas as pd
from pathlib import Path

def debug_section_detection(file_path, month_year="Apr 25"):
    """Debug section detection in detail"""
    print(f"\n=== DEBUGGING SECTION DETECTION ===")
    print(f"File: {file_path}")
    print(f"Target month: {month_year}")
    
    try:
        # Load workbook
        print("\n1. Loading workbook...")
        wb = load_wip_workbook(file_path)
        print(f"   Workbook loaded successfully")
        print(f"   Available sheets: {wb.sheetnames}")
        
        # Find/create monthly tab
        print(f"\n2. Finding monthly tab '{month_year}'...")
        ws = find_or_create_monthly_tab(wb, month_year)
        print(f"   Active sheet: {ws.title}")
        
        # Scan for section headers
        print(f"\n3. Scanning for section headers...")
        print(f"   Sheet has {ws.max_row} rows and {ws.max_column} columns")
        
        # Look for 5040 section
        print(f"\n4. Looking for '5040' section...")
        for row in range(1, min(50, ws.max_row + 1)):  # Check first 50 rows
            for col in range(1, min(10, ws.max_column + 1)):  # Check first 10 columns
                cell = ws.cell(row=row, column=col)
                if cell.value and '5040' in str(cell.value):
                    print(f"   Found '5040' at Row {row}, Col {col}: '{cell.value}'")
        
        # Look for 5030 section  
        print(f"\n5. Looking for '5030' section...")
        for row in range(1, min(100, ws.max_row + 1)):  # Check first 100 rows
            for col in range(1, min(10, ws.max_column + 1)):  # Check first 10 columns
                cell = ws.cell(row=row, column=col)
                if cell.value and '5030' in str(cell.value):
                    print(f"   Found '5030' at Row {row}, Col {col}: '{cell.value}'")
        
        # Test the actual find_section_markers function
        print(f"\n6. Testing find_section_markers function...")
        section_5040, section_5030 = find_section_markers(ws)
        print(f"   find_section_markers returned: 5040={section_5040}, 5030={section_5030}")
        
        # Check specific rows we expect
        print(f"\n7. Checking expected rows...")
        row_2_col_1 = ws.cell(row=2, column=1).value
        row_69_col_1 = ws.cell(row=69, column=1).value
        print(f"   Row 2, Col 1: '{row_2_col_1}'")
        print(f"   Row 69, Col 1: '{row_69_col_1}'")
        
        # Check row 2 across multiple columns
        print(f"\n8. Checking Row 2 across columns...")
        for col in range(1, 11):
            cell_value = ws.cell(row=2, column=col).value
            if cell_value:
                print(f"   Row 2, Col {col}: '{cell_value}'")
                
        # Check row 69 across multiple columns
        print(f"\n9. Checking Row 69 across columns...")
        for col in range(1, 11):
            cell_value = ws.cell(row=69, column=col).value
            if cell_value:
                print(f"   Row 69, Col {col}: '{cell_value}'")
        
        print(f"\n=== DEBUG COMPLETE ===")
        
    except Exception as e:
        print(f"ERROR during debug: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # Test with the test file
    test_file = "test_data/WIP Report.xlsm"
    if Path(test_file).exists():
        debug_section_detection(test_file)
    else:
        print(f"Test file not found: {test_file}")
        print("Available files in test_data:")
        test_dir = Path("test_data")
        if test_dir.exists():
            for file in test_dir.iterdir():
                if file.suffix in ['.xlsx', '.xlsm']:
                    print(f"  - {file.name}") 