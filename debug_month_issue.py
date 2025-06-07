#!/usr/bin/env python3

import sys
import os
from datetime import datetime
sys.path.append('/Users/richardrierson/Desktop/Projects/WIP/src')

from data_processing.excel_integration_v2 import find_or_create_monthly_tab, load_wip_workbook, find_section_markers

def test_different_months():
    """Test section detection across different monthly tabs"""
    
    # Find the WIP Report file
    wip_files = []
    test_data_dir = 'test_data'
    if os.path.exists(test_data_dir):
        for file in os.listdir(test_data_dir):
            if 'WIP Report' in file and file.endswith(('.xlsx', '.xlsm')) and not file.startswith('~$'):
                wip_files.append(os.path.join(test_data_dir, file))

    if not wip_files:
        print('No WIP Report files found')
        return
        
    file_path = wip_files[0]
    print(f'Testing month/year tabs in: {file_path}')
    
    # Load workbook
    wb = load_wip_workbook(file_path)
    print(f'Available sheets: {wb.sheetnames}')
    
    # Test current month (what Streamlit would default to)
    current_date = datetime.now()
    current_month_year = current_date.strftime("%b %y")
    print(f'\nCurrent month (Streamlit default): {current_month_year}')
    
    try:
        current_ws = find_or_create_monthly_tab(wb, current_month_year)
        current_markers = find_section_markers(current_ws, ["5040", "5030"])
        print(f'Current month sections: {current_markers}')
    except Exception as e:
        print(f'Error with current month: {str(e)}')
    
    # Test Apr 25 (the one that works)
    apr_month_year = "Apr 25"
    print(f'\nApril 2025: {apr_month_year}')
    
    try:
        apr_ws = find_or_create_monthly_tab(wb, apr_month_year)
        apr_markers = find_section_markers(apr_ws, ["5040", "5030"])
        print(f'April 2025 sections: {apr_markers}')
    except Exception as e:
        print(f'Error with April 2025: {str(e)}')
    
    # Test a few other months that exist
    test_months = ["Dec 24", "Nov 24", "Mar 25", "Feb 25"]
    
    for month_year in test_months:
        if month_year in wb.sheetnames:
            print(f'\nTesting {month_year}:')
            try:
                ws = find_or_create_monthly_tab(wb, month_year)
                markers = find_section_markers(ws, ["5040", "5030"])
                print(f'{month_year} sections: {markers}')
            except Exception as e:
                print(f'Error with {month_year}: {str(e)}')

if __name__ == "__main__":
    test_different_months() 