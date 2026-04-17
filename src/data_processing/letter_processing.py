"""
Letter Processing Module

Parses overdue invoice spreadsheets for the Lien Letter Generator.

Phase 2 format:
- Row 1 is the header row (no metadata to skip).
- 20 columns with a fixed positional layout (see LIEN_INVOICE_POSITIONAL).
- Duplicate header names (two "State", "Zip Code", "Suite", "City" columns each
  for Manager/Service/Owner sections) — positional mapping is authoritative.
"""

import pandas as pd
import logging
from typing import Tuple, List

from .column_mapping import (
    LIEN_INVOICE_POSITIONAL,
    map_columns_by_position,
    apply_column_mapping,
)

logger = logging.getLogger(__name__)

# Fields that must be present for a usable customer/manager letter
REQUIRED_FIELDS = [
    'customer_name', 'service_address', 'service_city',
    'service_state', 'service_zip', 'invoice_total'
]

# Fields that together make up a complete owner mailing
OWNER_REQUIRED_FIELDS = [
    'owner_name', 'owner_address', 'owner_city', 'owner_state', 'owner_zip'
]


def _is_blank(value) -> bool:
    """Return True for None, NaN, and empty/whitespace-only strings."""
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False


def _classify_owner(row: pd.Series) -> Tuple[bool, str]:
    """
    Decide whether an owner letter should be generated for a row.

    Returns:
        (owner_letter_needed, reason)
        reason is 'ok' | 'blank' | 'same_as_management' | 'incomplete:<details>'
    """
    name = row.get('owner_name')

    # "Same as Management" is handled before blank-name so it can't be
    # mistaken for partial data.
    if not _is_blank(name):
        name_lower = str(name).strip().lower()
        if name_lower.startswith('same'):
            return False, 'same_as_management'

    # Gather which of the five owner fields have values.
    present = [f for f in OWNER_REQUIRED_FIELDS if not _is_blank(row.get(f))]
    missing = [f for f in OWNER_REQUIRED_FIELDS if _is_blank(row.get(f))]

    # All five present → generate letter.
    if not missing:
        return True, 'ok'

    # None present → normal "no owner" row, no warning.
    if not present:
        return False, 'blank'

    # Partial data in any combination → warning, no letter.
    return False, f"incomplete:{','.join(present)}|missing:{','.join(missing)}"


def parse_invoice_spreadsheet(uploaded_file) -> Tuple[pd.DataFrame, List[str]]:
    """
    Parse a Phase 2 overdue invoice spreadsheet for lien letter generation.

    Args:
        uploaded_file: File-like object (BytesIO, Streamlit UploadedFile, or path)

    Returns:
        Tuple of (DataFrame with canonical columns, list of warning messages)
        DataFrame includes a computed `owner_letter_needed` boolean column.

    Raises:
        ValueError: If the spreadsheet has fewer than the expected columns or
                    no usable data rows.
    """
    warnings: List[str] = []

    # Row 1 is the header row in Phase 2.
    df = pd.read_excel(uploaded_file, header=0)
    logger.info(f"Read spreadsheet with {len(df)} data rows, {len(df.columns)} columns")

    if len(df.columns) < len(LIEN_INVOICE_POSITIONAL):
        raise ValueError(
            f"Expected at least {len(LIEN_INVOICE_POSITIONAL)} columns in the "
            f"spreadsheet (found {len(df.columns)}). Please use the Phase 2 "
            f"template with Manager / Customer / Invoice / Owner sections."
        )

    # Positional mapping is authoritative — duplicate headers ("State", "Zip
    # Code", "Suite", "City") can't be resolved by name alone.
    rename_map = map_columns_by_position(df)
    df = apply_column_mapping(df, rename_map)

    # Keep only the canonical columns (drop any trailing extras).
    canonical_cols = [c for c in LIEN_INVOICE_POSITIONAL if c in df.columns]
    df = df[canonical_cols].copy()

    # Drop rows where invoice_total is null/zero.
    df['invoice_total'] = pd.to_numeric(df['invoice_total'], errors='coerce')
    before = len(df)
    df = df[df['invoice_total'].notna() & (df['invoice_total'] != 0)].copy()
    dropped = before - len(df)
    if dropped:
        logger.info(f"Dropped {dropped} rows with null/zero invoice_total")

    if len(df) == 0:
        raise ValueError(
            "No invoice data found after filtering (all rows had zero or "
            "missing totals)."
        )

    # Strip whitespace on string-ish columns, convert 'nan' back to None.
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({'nan': None, 'None': None, '': None})

    # Normalize invoice_number to string, strip trailing .0
    df['invoice_number'] = (
        df['invoice_number']
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .replace({'nan': None, 'None': None})
    )

    # Normalize zip codes — strings, preserve leading zeros.
    for zip_col in ['service_zip', 'manager_zip', 'owner_zip']:
        if zip_col in df.columns:
            df[zip_col] = (
                df[zip_col]
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .replace({'nan': None, 'None': None})
            )

    # Reset index so row numbers line up with enumerate().
    df = df.reset_index(drop=True)

    # Owner logic + required-field warnings.
    owner_needed_flags: List[bool] = []
    for idx, row in df.iterrows():
        row_num = idx + 2  # spreadsheet row (header is row 1, data starts at 2)
        customer_name = row.get('customer_name') or 'Unknown'

        # Required-field warnings for customer/manager letters.
        missing = [
            f for f in REQUIRED_FIELDS
            if f not in df.columns or _is_blank(row.get(f))
        ]
        if missing:
            warnings.append(f"Row {row_num} ({customer_name}): missing {', '.join(missing)}")

        # Owner classification.
        needed, reason = _classify_owner(row)
        owner_needed_flags.append(needed)
        if reason.startswith('incomplete:'):
            # reason format: "incomplete:<present>|missing:<missing>"
            parts = reason.split('|')
            present = parts[0].split(':', 1)[1] or '(nothing)'
            missing_owner = parts[1].split(':', 1)[1] or '(nothing)'
            warnings.append(
                f"Row {row_num} ({customer_name}): Owner data is incomplete — "
                f"has {present} but missing {missing_owner}. No owner letter "
                f"generated. Please verify."
            )

    df['owner_letter_needed'] = owner_needed_flags

    logger.info(
        f"Parsed {len(df)} invoice rows, "
        f"{sum(owner_needed_flags)} owner letters needed, "
        f"{len(warnings)} warnings"
    )
    return df, warnings
