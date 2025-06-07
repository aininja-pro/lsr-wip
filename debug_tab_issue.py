#!/usr/bin/env python3
import sys
sys.path.append('/app/src')
from data_processing.excel_integration_v2 import find_or_create_monthly_tab, load_wip_workbook, find_section_markers
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Load the workbook
print("Loading workbook...")
wb = load_wip_workbook('/app/test_data/Master WIP Report.xlsx')
print(f"Available sheets: {wb.sheetnames}")

# Test different month formats
test_formats = ["Apr 25", "April 25", "April 2025", "Apr 2025"]

for month_format in test_formats:
    print(f"\n=== Testing month format: '{month_format}' ===")
    try:
        ws = find_or_create_monthly_tab(wb, month_format)
        print(f"Found/created tab: '{ws.title}'")
        
        # Test section detection
        sections = find_section_markers(ws, ['5040', '5030'])
        print(f"Section detection results: {sections}")
        
        # Check specific cells we know should have the headers
        if ws.title in ['April 25', 'Apr 25']:
            print("Checking known header cells:")
            cell_2_2 = ws.cell(row=2, column=2)
            print(f"Row 2, Col 2: '{cell_2_2.value}'")
            
            cell_69_2 = ws.cell(row=69, column=2)
            print(f"Row 69, Col 2: '{cell_69_2.value}'")
            
    except Exception as e:
        print(f"Error with format '{month_format}': {e}")

print("\n=== Final Recommendation ===")
print("The issue seems to be in tab name matching or section detection.")
print("Check the above results to see which format works correctly.") 