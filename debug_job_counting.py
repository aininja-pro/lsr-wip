#!/usr/bin/env python3
"""
Debug script to understand why job counting is returning 0 
when there should be jobs in the sections.
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

def debug_section_content(worksheet, section_name, start_row, start_col, rows_to_check=20):
    """Debug what's actually in a section."""
    print(f"\n🔍 Debugging {section_name} section content starting at Row {start_row}, Col {start_col}")
    print("Row".ljust(6) + "Col A".ljust(20) + "Col B".ljust(20) + "Col C".ljust(20) + "Col D".ljust(20))
    print("-" * 80)
    
    for row_offset in range(rows_to_check):
        row_num = start_row + row_offset
        
        # Get values from columns A, B, C, D
        col_a = worksheet.cell(row=row_num, column=1).value
        col_b = worksheet.cell(row=row_num, column=2).value
        col_c = worksheet.cell(row=row_num, column=3).value
        col_d = worksheet.cell(row=row_num, column=4).value
        
        # Format for display
        col_a_str = str(col_a)[:18] if col_a else ""
        col_b_str = str(col_b)[:18] if col_b else ""
        col_c_str = str(col_c)[:18] if col_c else ""
        col_d_str = str(col_d)[:18] if col_d else ""
        
        print(f"{row_num:<6}{col_a_str:<20}{col_b_str:<20}{col_c_str:<20}{col_d_str:<20}")
        
        # Stop early if we hit several empty rows
        if not any([col_a, col_b, col_c, col_d]):
            empty_count = 0
            for check_offset in range(3):
                check_row = row_num + check_offset
                if not any([
                    worksheet.cell(row=check_row, column=1).value,
                    worksheet.cell(row=check_row, column=2).value,
                    worksheet.cell(row=check_row, column=3).value,
                    worksheet.cell(row=check_row, column=4).value
                ]):
                    empty_count += 1
            if empty_count >= 3:
                print("   [Stopping - found 3+ consecutive empty rows]")
                break

def debug_job_counting():
    """Debug job counting for April 25 tab specifically."""
    print("🔧 Debugging Job Counting in April 25 Tab")
    print("=" * 50)
    
    # Load the Master WIP Report
    master_report_file = Path("test_data") / "Master WIP Report.xlsx"
    
    if not master_report_file.exists():
        print(f"❌ Master WIP Report not found: {master_report_file}")
        return False
    
    try:
        # Load workbook
        workbook = load_wip_workbook(str(master_report_file))
        
        # Focus on April 25 tab (where user confirmed job counts)
        if 'April 25' not in workbook.sheetnames:
            print("❌ April 25 tab not found")
            return False
        
        worksheet = workbook['April 25']
        print("✅ Loaded April 25 worksheet")
        
        # Find section markers
        markers = find_section_markers(worksheet, ['5040', '5030'])
        print(f"Section markers: {markers}")
        
        # Debug 5040 section content
        if markers['5040']:
            start_row, start_col = markers['5040']
            print(f"\n📊 5040 SECTION DEBUG")
            print(f"Header found at Row {start_row}, Col {start_col}")
            
            # Show header row
            header_cell = worksheet.cell(row=start_row, column=start_col)
            print(f"Header cell content: '{header_cell.value}'")
            
            # Show content starting from next row
            debug_section_content(worksheet, "5040", start_row + 1, start_col, 15)
            
            # Manual job count with different criteria
            print(f"\n🔢 Manual job counting (5040 section):")
            job_count = 0
            for row_offset in range(20):
                row_num = start_row + 1 + row_offset
                cell = worksheet.cell(row=row_num, column=1)  # Column A
                
                if cell.value:
                    value = str(cell.value).strip()
                    print(f"   Row {row_num}: '{value}'")
                    
                    # More lenient job detection
                    if value and len(value) > 2 and not value.lower().startswith('total'):
                        job_count += 1
                        print(f"      ✅ Counted as job #{job_count}")
                    else:
                        print(f"      ❌ Skipped (header/total/empty)")
                elif row_offset > 5:  # Only break after checking a few rows
                    print(f"   Row {row_num}: [Empty - stopping count]")
                    break
            
            print(f"   Manual count result: {job_count} jobs")
        
        # Debug 5030 section content
        if markers['5030']:
            start_row, start_col = markers['5030']
            print(f"\n📊 5030 SECTION DEBUG")
            print(f"Header found at Row {start_row}, Col {start_col}")
            
            # Show header row
            header_cell = worksheet.cell(row=start_row, column=start_col)
            print(f"Header cell content: '{header_cell.value}'")
            
            # Show content starting from next row
            debug_section_content(worksheet, "5030", start_row + 1, start_col, 15)
            
            # Manual job count with different criteria
            print(f"\n🔢 Manual job counting (5030 section):")
            job_count = 0
            for row_offset in range(25):  # Check more rows for 5030
                row_num = start_row + 1 + row_offset
                cell = worksheet.cell(row=row_num, column=1)  # Column A (descriptions)
                
                if cell.value:
                    value = str(cell.value).strip()
                    print(f"   Row {row_num}: '{value}'")
                    
                    # More lenient job detection for descriptions
                    if value and len(value) > 2 and not value.lower().startswith('total'):
                        job_count += 1
                        print(f"      ✅ Counted as job #{job_count}")
                    else:
                        print(f"      ❌ Skipped (header/total/empty)")
                elif row_offset > 8:  # Only break after checking several rows
                    print(f"   Row {row_num}: [Empty - stopping count]")
                    break
            
            print(f"   Manual count result: {job_count} jobs")
        
        return True
        
    except Exception as e:
        print(f"❌ ERROR during debugging: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    debug_job_counting() 