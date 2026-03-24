"""
Letter Processing Module

Parses overdue invoice spreadsheets for the Lien Letter Generator.
Handles metadata row skipping, fuzzy column matching, and data normalization.
"""

import pandas as pd
import logging
from typing import Tuple, List
from .column_mapping import COLUMN_MAPPINGS, map_columns_for_file_type, apply_column_mapping

logger = logging.getLogger(__name__)

# Fields that must be present for a usable letter
REQUIRED_FIELDS = [
    'customer_name', 'service_address', 'service_city',
    'service_state', 'service_zip', 'invoice_total'
]


def _find_header_row(df_raw: pd.DataFrame) -> int:
    """
    Scan the first 10 rows to find the header row.

    The header row is the first row where 3+ cells match known
    lien_invoice column names from COLUMN_MAPPINGS.

    Returns the 0-based row index, or raises ValueError if not found.
    """
    known_names = set()
    for variations in COLUMN_MAPPINGS['lien_invoice'].values():
        for v in variations:
            known_names.add(v.lower())

    for row_idx in range(min(10, len(df_raw))):
        row_values = df_raw.iloc[row_idx]
        matches = sum(
            1 for val in row_values
            if isinstance(val, str) and val.strip().lower() in known_names
        )
        if matches >= 3:
            logger.info(f"Found header row at index {row_idx} ({matches} column matches)")
            return row_idx

    raise ValueError(
        "Could not find column headers. Expected columns: "
        "Customer Name, Invoice Total, Service Location Address 1, etc."
    )


def parse_invoice_spreadsheet(uploaded_file) -> Tuple[pd.DataFrame, List[str]]:
    """
    Parse an overdue invoice spreadsheet for lien letter generation.

    Args:
        uploaded_file: File-like object (BytesIO, Streamlit UploadedFile, or file path)

    Returns:
        Tuple of (DataFrame with normalized columns, list of warning messages)

    Raises:
        ValueError: If headers can't be found or no data rows exist
    """
    warnings = []

    # Read all rows without assuming a header row
    df_raw = pd.read_excel(uploaded_file, header=None)

    # Find the header row by scanning for known column names
    header_idx = _find_header_row(df_raw)

    # Re-read with the correct header row
    df = pd.read_excel(uploaded_file, header=header_idx)
    logger.info(f"Read spreadsheet with {len(df)} data rows, columns: {list(df.columns)}")

    if len(df) == 0:
        raise ValueError("No invoice data found in spreadsheet")

    # Apply column matching to normalize names (strict mode — no fuzzy matching,
    # since fuzzy matching can mismap e.g. "Invoice Date" to "invoice_number")
    column_mapping = map_columns_for_file_type(list(df.columns), 'lien_invoice', strict_mode=True)
    df = apply_column_mapping(df, column_mapping)

    mapped_cols = set(column_mapping.values())
    logger.info(f"Mapped columns: {mapped_cols}")

    # Drop rows where invoice_total is null or zero
    if 'invoice_total' in df.columns:
        df['invoice_total'] = pd.to_numeric(df['invoice_total'], errors='coerce')
        before = len(df)
        df = df[df['invoice_total'].notna() & (df['invoice_total'] != 0)]
        dropped = before - len(df)
        if dropped:
            logger.info(f"Dropped {dropped} rows with null/zero invoice_total")

    if len(df) == 0:
        raise ValueError("No invoice data found after filtering (all rows had zero or missing totals)")

    # Strip whitespace from string columns
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            # Convert 'nan' strings back to None
            df[col] = df[col].replace('nan', None)

    # Normalize invoice_number to string, strip trailing .0
    if 'invoice_number' in df.columns:
        df['invoice_number'] = (
            df['invoice_number']
            .astype(str)
            .str.replace(r'\.0$', '', regex=True)
            .replace('nan', None)
            .replace('None', None)
        )

    # Normalize zip codes — ensure strings, preserve leading zeros
    for zip_col in ['service_zip', 'bill_to_zip']:
        if zip_col in df.columns:
            df[zip_col] = (
                df[zip_col]
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .replace('nan', None)
                .replace('None', None)
            )

    # Set owner to None if column is missing
    if 'owner' not in df.columns:
        df['owner'] = None

    # Generate warnings for rows missing required fields
    df = df.reset_index(drop=True)
    for idx, row in df.iterrows():
        row_num = idx + header_idx + 2  # Convert to 1-based spreadsheet row number
        missing = []
        for field in REQUIRED_FIELDS:
            if field not in df.columns:
                missing.append(field)
            elif row[field] is None or (isinstance(row[field], str) and row[field].strip() == ''):
                missing.append(field)
            elif field == 'invoice_total' and pd.isna(row[field]):
                missing.append(field)
        if missing:
            warnings.append(f"Row {row_num}: missing {', '.join(missing)}")

    logger.info(f"Parsed {len(df)} invoice rows with {len(warnings)} warnings")
    return df, warnings
