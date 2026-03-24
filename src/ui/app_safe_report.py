#!/usr/bin/env python3
"""
Safe WIP Report Automation - Report Generation Only
Generates update reports that can be manually copied into Excel
This avoids ALL Excel corruption issues
"""

import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime, date, timedelta
from pathlib import Path
import logging

# Import our data processing functions
import sys
sys.path.append(str(Path(__file__).parent.parent))

from data_processing.aggregation import (
    filter_gl_accounts,
    compute_amounts,
    aggregate_gl_data
)
from data_processing.merge_data import merge_wip_with_gl
from data_processing.column_mapping import map_dataframe_columns
from data_processing.letter_processing import parse_invoice_spreadsheet
from letter_generation.letter_template import generate_letter, make_letter_filename

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def initialize_session_state():
    """Initialize session state variables"""
    defaults = {
        # WIP Reports state
        'files_uploaded': {},
        'merged_data': None,
        'results_ready': False,
        'labor_df': None,
        'material_df': None,
        'excel_report': None,
        'month_year': None,
        'gl_entries': 0,
        # Lien Letters state
        'lien_parsed_df': None,
        'lien_parse_warnings': [],
        'lien_generated_zip': None,
        'lien_generation_complete': False,
        'lien_summary_df': None,
    }
    for key, default in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default

def process_data(wip_bytes, gl_bytes, include_closed):
    """Process the data using our existing functions"""
    try:
        with st.spinner("Processing GL data..."):
            gl_df = pd.read_excel(io.BytesIO(gl_bytes))
            logger.info(f"Available GL Inquiry columns: {list(gl_df.columns)}")
            gl_df = map_dataframe_columns(gl_df, 'gl_inquiry')

            filtered_gl = filter_gl_accounts(gl_df)
            amounts_gl = compute_amounts(filtered_gl)
            gl_summary = aggregate_gl_data(amounts_gl)
            st.session_state.gl_entries = len(gl_summary)

        with st.spinner("Merging data..."):
            wip_df = pd.read_excel(io.BytesIO(wip_bytes))
            logger.info(f"Available WIP Worksheet columns: {list(wip_df.columns)}")
            wip_df = map_dataframe_columns(wip_df, 'wip_worksheet')
            logger.info(f"WIP Worksheet columns after mapping: {list(wip_df.columns)}")

            merged_df = merge_wip_with_gl(wip_df, gl_summary, include_closed)

        return merged_df

    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        logger.error(f"Processing error: {e}")
        return None

def generate_update_reports(merged_df):
    """Generate reports with EXACTLY the fields requested"""
    
    # 5040 Section - Labor Report (with Percent Complete column)
    labor_data = []
    for _, job in merged_df.iterrows():
        labor_actual = job.get('5040', 0) or job.get('Labor Actual', 0) or job.get('Sub Labor', 0)
        estimated_labor = job.get('Total Subcontract Est', 0)
        
        # Calculate percent complete (avoid division by zero, cap at 100%)
        if estimated_labor > 0:
            percent_complete = min((labor_actual / estimated_labor) * 100, 100.0)
        else:
            percent_complete = 0.0
        
        labor_data.append({
            'Job Number': job.get('Job Number', ''),
            'Job Description': job.get('Job Name', job.get('Job Description', '')),
            'Contract Amount': job.get('Original Contract Amount', 0),  # Using actual column name
            'Estimated Sub Labor Costs': estimated_labor,  # Using actual column name
            'Monthly Sub Labor Costs': labor_actual,
            'Percent Complete': percent_complete,  # New column
            'Amount Billed': job.get('Amount Billed', 0)  # Using properly calculated Amount Billed from GL aggregation
        })
    
    labor_df = pd.DataFrame(labor_data)
    
    # Convert to numeric and filter to include jobs with labor costs OR billing
    labor_df['Monthly Sub Labor Costs'] = pd.to_numeric(labor_df['Monthly Sub Labor Costs'], errors='coerce').fillna(0)
    labor_df['Amount Billed'] = pd.to_numeric(labor_df['Amount Billed'], errors='coerce').fillna(0)
    
    # Include jobs that have either labor costs (!=0) OR have been billed
    labor_df = labor_df[(labor_df['Monthly Sub Labor Costs'] != 0) | (labor_df['Amount Billed'] > 0)]
    
    # 5030 Section - Material Report (4 fields only)
    material_data = []
    for _, job in merged_df.iterrows():
        material_actual = job.get('5030', 0) or job.get('Material Actual', 0) or job.get('Material', 0)
        
        material_data.append({
            'Job Number': job.get('Job Number', ''),
            'Job Description': job.get('Job Name', job.get('Job Description', '')),
            'Estimated Material Costs': job.get('Total Material Estimate', 0),  # Using actual column name
            'Monthly Material Costs': material_actual
        })
    
    material_df = pd.DataFrame(material_data)
    
    # Convert to numeric and filter out rows where Monthly Material Costs is 0 or blank (include negative values)
    material_df['Monthly Material Costs'] = pd.to_numeric(material_df['Monthly Material Costs'], errors='coerce').fillna(0)
    material_df = material_df[material_df['Monthly Material Costs'] != 0]
    
    return labor_df, material_df

def _auto_adjust_column_widths(worksheet):
    """Auto-adjust column widths based on cell content."""
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except (TypeError, AttributeError):
                pass
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)


def _apply_number_format(worksheet, col_names, fmt):
    """Apply a number format to named columns in a worksheet."""
    headers = [cell.value for cell in worksheet[1]]
    for col_name in col_names:
        if col_name in headers:
            col_index = headers.index(col_name) + 1
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=col_index)
                if cell.value is not None and isinstance(cell.value, (int, float)):
                    cell.number_format = fmt


def create_excel_update_report(labor_df, material_df):
    """Create a comprehensive Excel report with all updates"""

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        labor_df.to_excel(writer, sheet_name='5040_Labor_Updates', index=False)
        material_df.to_excel(writer, sheet_name='5030_Material_Updates', index=False)

        # Format data sheets
        currency_columns = {
            '5040_Labor_Updates': ['Contract Amount', 'Monthly Sub Labor Costs', 'Estimated Sub Labor Costs', 'Amount Billed'],
            '5030_Material_Updates': ['Monthly Material Costs', 'Estimated Material Costs']
        }

        for sheet_name in ['5040_Labor_Updates', '5030_Material_Updates']:
            worksheet = writer.sheets[sheet_name]
            _apply_number_format(worksheet, currency_columns[sheet_name], '$#,##0.00')

            # Percent Complete needs value conversion (already *100) before formatting
            if sheet_name == '5040_Labor_Updates':
                headers = [cell.value for cell in worksheet[1]]
                if 'Percent Complete' in headers:
                    col_index = headers.index('Percent Complete') + 1
                    for row in range(2, worksheet.max_row + 1):
                        cell = worksheet.cell(row=row, column=col_index)
                        if cell.value is not None and isinstance(cell.value, (int, float)):
                            cell.value = cell.value / 100
                            cell.number_format = '0.00%'

            _auto_adjust_column_widths(worksheet)

        # Summary sheet
        labor_actual_total = labor_df['Monthly Sub Labor Costs'].sum()
        material_actual_total = material_df['Monthly Material Costs'].sum()
        labor_budget_total = labor_df['Estimated Sub Labor Costs'].sum()
        material_budget_total = material_df['Estimated Material Costs'].sum()
        labor_variance = labor_actual_total - labor_budget_total
        material_variance = material_actual_total - material_budget_total
        contract_total = labor_df['Contract Amount'].sum()
        billed_total = labor_df['Amount Billed'].sum()

        summary_data = {
            'Section': ['5040 - Labor', '5030 - Material', 'Total'],
            'Jobs Count': [len(labor_df), len(material_df), len(labor_df)],
            'Total Contract Amount': [contract_total, 0, contract_total],
            'Total Actual': [labor_actual_total, material_actual_total, labor_actual_total + material_actual_total],
            'Total Budget': [labor_budget_total, material_budget_total, labor_budget_total + material_budget_total],
            'Total Variance': [labor_variance, material_variance, labor_variance + material_variance],
            'Total Amount Billed': [billed_total, 0, billed_total]
        }

        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        summary_sheet = writer.sheets['Summary']
        _apply_number_format(summary_sheet, ['Total Contract Amount', 'Total Actual', 'Total Budget', 'Total Variance', 'Total Amount Billed'], '$#,##0.00')
        _auto_adjust_column_widths(summary_sheet)

        # Instructions sheet
        instructions = [
            "WIP REPORT UPDATE INSTRUCTIONS",
            "",
            "This report contains all the updates for your WIP Report without modifying the original file.",
            "This approach preserves ALL formulas, formatting, and macros in your Excel file.",
            "",
            "HOW TO USE:",
            "",
            "1. LABOR SECTION (5040):",
            "   - Open the '5040_Labor_Updates' tab in this report",
            "   - Copy the 'Monthly Sub Labor Costs' column values",
            "   - Paste them into the appropriate column in your WIP Report's 5040 section",
            "",
            "2. MATERIAL SECTION (5030):",
            "   - Open the '5030_Material_Updates' tab in this report",
            "   - Copy the 'Monthly Material Costs' column values",
            "   - Paste them into the appropriate column in your WIP Report's 5030 section",
            "",
            "3. VERIFICATION:",
            "   - Check the 'Summary' tab for totals and variance analysis",
            "   - Variances > $1,000 should be reviewed",
            "",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ]

        instructions_df = pd.DataFrame({'Instructions': instructions})
        instructions_df.to_excel(writer, sheet_name='Instructions', index=False)
        writer.sheets['Instructions'].column_dimensions['A'].width = 80

    buffer.seek(0)
    return buffer.getvalue()

def display_file_upload_section():
    """Display file upload interface"""
    st.markdown("#### 📁 File Upload")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**WIP Worksheet Export**")
        wip_file = st.file_uploader(
            "Upload WIP Worksheet",
            type=['xlsx'],
            key='wip_worksheet'
        )
        if wip_file:
            st.session_state.files_uploaded['wip'] = wip_file.getvalue()
            st.success(f"✅ {wip_file.name}")
    
    with col2:
        st.markdown("**GL Inquiry Export**")
        gl_file = st.file_uploader(
            "Upload GL Inquiry",
            type=['xlsx'],
            key='gl_inquiry'
        )
        if gl_file:
            st.session_state.files_uploaded['gl'] = gl_file.getvalue()
            st.success(f"✅ {gl_file.name}")

def display_sidebar_options():
    """Display processing options in sidebar"""
    st.sidebar.markdown("### ⚙️ Options")
    
    # Month/Year selector with proper dropdown
    st.sidebar.markdown("**Report Period**")
    
    current_year = datetime.now().year
    years = list(range(current_year - 2, current_year + 2))
    months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ]
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        selected_month = st.selectbox("Month", months, index=3)  # Default to Apr
    with col2:
        selected_year = st.selectbox("Year", years, index=len(years)//2)
    
    # Format as MMM YY
    month_year = f"{selected_month} {str(selected_year)[-2:]}"
    
    st.sidebar.markdown("**Processing Settings**")
    include_closed = st.sidebar.checkbox(
        "Include Closed Jobs", 
        value=False,
        help="Check this to include jobs with 'Closed' status in the report. Useful for quarterly reviews."
    )
    
    return include_closed, month_year

def render_wip_reports():
    """Render the WIP Reports tool — pure extraction of existing logic."""
    st.markdown("### WIP Report Generator")
    st.caption("Generate update reports without modifying Excel files")

    # Sidebar options
    include_closed, month_year = display_sidebar_options()

    # File Upload Section
    display_file_upload_section()

    st.markdown("---")

    # Process Button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("Generate Update Reports", type="primary", use_container_width=True):
            if len(st.session_state.files_uploaded) >= 2:
                merged_df = process_data(
                    st.session_state.files_uploaded['wip'],
                    st.session_state.files_uploaded['gl'],
                    include_closed
                )

                if merged_df is not None:
                    st.session_state.merged_data = merged_df

                    with st.spinner("Generating update reports..."):
                        labor_df, material_df = generate_update_reports(merged_df)
                        excel_report = create_excel_update_report(labor_df, material_df)

                    st.session_state.results_ready = True
                    st.session_state.labor_df = labor_df
                    st.session_state.material_df = material_df
                    st.session_state.excel_report = excel_report
                    st.session_state.month_year = month_year

            else:
                st.error("Please upload both the WIP Worksheet and GL Inquiry files")

    # Download button
    if st.session_state.get('results_ready', False):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.download_button(
                label="Download Update Reports (Excel)",
                data=st.session_state.excel_report,
                file_name=f"WIP_Update_Reports_{st.session_state.month_year.replace(' ', '')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download comprehensive update reports for your WIP Excel file",
                use_container_width=True
            )

    # Results Section
    if st.session_state.get('results_ready', False):
        st.markdown("---")

        col1, col2 = st.columns([1, 1])

        with col1:
            st.markdown("#### Processing Status")
            st.success(f"Processed {st.session_state.gl_entries} GL entries")
            st.success(f"Merged data for {len(st.session_state.merged_data)} jobs")
            st.success("Update reports generated successfully!")

        with col2:
            st.markdown("#### Report Summary")

            labor_variance = st.session_state.labor_df['Monthly Sub Labor Costs'].sum() - st.session_state.labor_df['Estimated Sub Labor Costs'].sum()
            material_variance = st.session_state.material_df['Monthly Material Costs'].sum() - st.session_state.material_df['Estimated Material Costs'].sum()
            total_variance = labor_variance + material_variance

            summary_data = {
                'Category': ['Jobs Processed', 'Labor Actual', 'Material Actual', 'Labor Variance', 'Material Variance', 'Total Variance'],
                'Value': [
                    f"{len(st.session_state.merged_data)} jobs",
                    f"${st.session_state.labor_df['Monthly Sub Labor Costs'].sum():,.2f}",
                    f"${st.session_state.material_df['Monthly Material Costs'].sum():,.2f}",
                    f"${labor_variance:,.2f}",
                    f"${material_variance:,.2f}",
                    f"${total_variance:,.2f}"
                ]
            }

            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("#### Data Preview")

        tab1, tab2 = st.tabs(["5040 - Labor Updates", "5030 - Material Updates"])

        with tab1:
            st.markdown("**Labor Section Data (Non-Zero Values Only)**")
            if len(st.session_state.labor_df) > 0:
                st.dataframe(st.session_state.labor_df, use_container_width=True, height=450)
            else:
                st.info("No labor entries with non-zero values found.")

        with tab2:
            st.markdown("**Material Section Data (Non-Zero Values Only)**")
            if len(st.session_state.material_df) > 0:
                st.dataframe(st.session_state.material_df, use_container_width=True, height=450)
            else:
                st.info("No material entries with non-zero values found.")


def _clear_lien_state():
    """Clear lien letter session state for a fresh upload."""
    st.session_state.lien_parsed_df = None
    st.session_state.lien_parse_warnings = []
    st.session_state.lien_generated_zip = None
    st.session_state.lien_generation_complete = False
    st.session_state.lien_summary_df = None


def render_lien_letters():
    """Render the Lien Letter Generator tool."""
    st.markdown("### Lien Letter Generator")
    st.caption("Generate pre-lien notice letters from an invoice spreadsheet")


    # Upload section
    st.markdown("**Upload Invoice Spreadsheet**")
    invoice_file = st.file_uploader(
        "Upload overdue invoice export (.xlsx)",
        type=['xlsx'],
        key='lien_invoice_upload'
    )

    # Parse on upload
    if invoice_file:
        file_bytes = invoice_file.getvalue()
        # Re-parse if it's a new file
        file_id = f"{invoice_file.name}_{len(file_bytes)}"
        if st.session_state.get('_lien_file_id') != file_id:
            _clear_lien_state()
            st.session_state._lien_file_id = file_id
            try:
                parsed_df, warnings = parse_invoice_spreadsheet(io.BytesIO(file_bytes))
                st.session_state.lien_parsed_df = parsed_df
                st.session_state.lien_parse_warnings = warnings
                st.success(f"{len(parsed_df)} invoices found")
            except ValueError as e:
                st.error(str(e))
                return

    parsed_df = st.session_state.get('lien_parsed_df')
    if parsed_df is None:
        return

    # Show warnings
    warnings = st.session_state.get('lien_parse_warnings', [])
    if warnings:
        with st.expander(f"Warnings ({len(warnings)})", expanded=False):
            for w in warnings:
                st.warning(w)

    # Settings
    st.markdown("---")
    st.markdown("**Settings**")
    col1, col2 = st.columns(2)
    with col1:
        deadline_days = st.number_input(
            "Payment deadline (days from today)",
            min_value=7, max_value=90, value=23, step=1,
            key='lien_deadline_days'
        )
    with col2:
        letter_date = date.today()
        deadline_date = letter_date + timedelta(days=deadline_days)
        st.markdown(f"**Letter date:** {letter_date.strftime('%B %d, %Y')}")
        st.markdown(f"**Deadline:** {deadline_date.strftime('%B %d, %Y')}")

    st.markdown("---")

    # Generate button
    logo_path = str(Path(__file__).parent.parent / 'assets' / 'lsr_logo.png')

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("Generate Letters", type="primary", use_container_width=True):
            summary_rows = []
            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                progress = st.progress(0)
                total = len(parsed_df)

                for idx, (_, row) in enumerate(parsed_df.iterrows()):
                    row_dict = row.to_dict()
                    customer = row_dict.get('customer_name') or 'Unknown'
                    inv_num = row_dict.get('invoice_number') or 'NO_INV'
                    inv_total = float(row_dict.get('invoice_total') or 0)
                    address_parts = [
                        row_dict.get('service_address') or '',
                        row_dict.get('service_city') or '',
                        row_dict.get('service_state') or '',
                    ]
                    address_str = ', '.join(p for p in address_parts if p)

                    status = "Generated"
                    try:
                        docx_bytes = generate_letter(row_dict, letter_date, deadline_date, logo_path)
                        docx_filename = make_letter_filename(customer, inv_num)
                        zf.writestr(f"docx/{docx_filename}", docx_bytes.getvalue())

                        # Try PDF conversion
                    except Exception as e:
                        status = f"Error: {e}"
                        logger.error(f"Letter generation failed for row {idx+1}: {e}")

                    summary_rows.append({
                        '#': idx + 1,
                        'Property': customer,
                        'Address': address_str,
                        'Invoice #': inv_num,
                        'Amount': f"${inv_total:,.2f}",
                        'Status': status,
                    })

                    progress.progress((idx + 1) / total)

            zip_buffer.seek(0)
            zip_bytes = zip_buffer.getvalue()
            st.session_state.lien_generated_zip = zip_bytes
            st.session_state.lien_zip_filename = f"Lien_Letters_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            st.session_state.lien_generation_complete = True
            st.session_state.lien_summary_df = pd.DataFrame(summary_rows)
            logger.info(f"ZIP generated: {len(zip_bytes)} bytes")

    # Results
    if st.session_state.get('lien_generation_complete', False):
        summary_df = st.session_state.lien_summary_df
        n_success = len(summary_df[summary_df['Status'] == 'Generated'])
        n_errors = len(summary_df[summary_df['Status'] != 'Generated'])
        total_amount = parsed_df['invoice_total'].sum()

        st.markdown("---")
        st.success(f"Generated {len(summary_df)} letters ({n_success} successful, {n_errors} errors)")
        st.markdown(f"**Total amount:** ${total_amount:,.2f}")
        st.markdown(f"**Payment deadline:** {deadline_date.strftime('%B %d, %Y')}")

        if n_errors > 0:
            st.warning(f"{n_errors} letters had issues — see Status column below")

        # Download button — placed before table so it's immediately visible
        zip_data = st.session_state.get('lien_generated_zip')
        if zip_data:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    label="Download All Letters (ZIP)",
                    data=zip_data,
                    file_name=st.session_state.get('lien_zip_filename', 'Lien_Letters.zip'),
                    mime="application/zip",
                    use_container_width=True,
                    key="lien_download_zip"
                )

        # Summary table
        st.markdown("#### Letter Summary")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        st.caption(f"Total: {len(summary_df)} letters, ${total_amount:,.2f}")


def main():
    st.set_page_config(
        page_title="LSR_Pay Tools",
        page_icon="📊",
        layout="wide"
    )

    # Apply custom theme
    from ui.styles import inject_custom_css
    inject_custom_css()

    initialize_session_state()

    # App header
    st.markdown("# LSR_Pay Tools")

    # Top-level navigation
    selected_tab = st.tabs(["WIP Reports", "Lien Letters"])

    with selected_tab[0]:
        render_wip_reports()

    with selected_tab[1]:
        render_lien_letters()


if __name__ == "__main__":
    main() 