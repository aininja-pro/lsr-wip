"""
Column Mapping Utility Module

This module provides standardized column mapping functionality to handle
variations in column headers across different Excel files from Sage exports
and WIP worksheets.
"""

import logging
from typing import Dict, List, Optional, Tuple
from difflib import SequenceMatcher


# Standard column mappings for different file types
COLUMN_MAPPINGS = {
    'gl_inquiry': {
        'Account': ['Account', 'Account Number', 'Acct', 'GL Account', 'Account Code'],
        'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number', 'Project No', 'Project'],
        'Debit': ['Debit', 'Debit Amount', 'DR', 'Dr', 'Debit Amt'],
        'Credit': ['Credit', 'Credit Amount', 'CR', 'Cr', 'Credit Amt'],
        'Description': ['Description', 'Desc', 'Transaction Description', 'GL Description'],
        'Date': ['Date', 'Transaction Date', 'GL Date', 'Post Date']
    },
    'wip_worksheet': {
        'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number', 'Project No', 'Project'],
        'Status': ['Status', 'Job Status', 'Project Status', 'State', 'Job State'],
        'Job Name': ['Job Name', 'Project Name', 'Description', 'Job Description', 'Project Description'],
        'Budget Material': ['Budget Material', 'Material Budget', 'Mat Budget', 'Budget Mat', 'Material Budgeted'],
        'Budget Labor': ['Budget Labor', 'Labor Budget', 'Lab Budget', 'Budget Lab', 'Labor Budgeted'],
        'Actual Material': ['Actual Material', 'Material Actual', 'Mat Actual', 'Actual Mat', 'Material To Date'],
        'Actual Labor': ['Actual Labor', 'Labor Actual', 'Lab Actual', 'Actual Lab', 'Labor To Date'],
        'Contract Amount': ['Contract Amount', 'Contract Value', 'Total Contract', 'Contract Total'],
        'Percent Complete': ['Percent Complete', '% Complete', 'Completion %', 'Progress %']
    },
    'wip_report': {
        'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number', 'Project No'],
        'Job Name': ['Job Name', 'Project Name', 'Description', 'Job Description'],
        'Material': ['Material', 'Materials', 'Mat', 'Material Cost'],
        'Labor': ['Labor', 'Labour', 'Lab', 'Labor Cost'],
        'Other': ['Other', 'Other Costs', 'Misc', 'Miscellaneous'],
        'Total': ['Total', 'Total Cost', 'Grand Total', 'Sum']
    }
}


def get_similarity_score(str1: str, str2: str) -> float:
    """
    Calculate similarity score between two strings using SequenceMatcher.
    
    Args:
        str1 (str): First string
        str2 (str): Second string
        
    Returns:
        float: Similarity score between 0 and 1
    """
    return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()


def find_best_column_match(target_column: str, available_columns: List[str], 
                          threshold: float = 0.6) -> Optional[str]:
    """
    Find the best matching column from available columns using fuzzy matching.
    
    Args:
        target_column (str): The target column name to match
        available_columns (List[str]): List of available column names
        threshold (float): Minimum similarity threshold (default: 0.6)
        
    Returns:
        Optional[str]: Best matching column name or None if no good match found
    """
    best_match = None
    best_score = 0.0
    
    for available_col in available_columns:
        score = get_similarity_score(target_column, available_col)
        if score > best_score and score >= threshold:
            best_score = score
            best_match = available_col
    
    logging.debug(f"Best match for '{target_column}': '{best_match}' (score: {best_score:.2f})")
    return best_match


def map_columns_for_file_type(available_columns: List[str], file_type: str,
                             strict_mode: bool = False) -> Dict[str, str]:
    """
    Map available columns to standard column names for a specific file type.
    
    Args:
        available_columns (List[str]): List of available column names in the file
        file_type (str): Type of file ('gl_inquiry', 'wip_worksheet', 'wip_report')
        strict_mode (bool): If True, only use exact matches from COLUMN_MAPPINGS
        
    Returns:
        Dict[str, str]: Mapping from available column names to standard names
        
    Raises:
        ValueError: If file_type is not recognized
    """
    if file_type not in COLUMN_MAPPINGS:
        raise ValueError(f"Unknown file type: {file_type}. Available types: {list(COLUMN_MAPPINGS.keys())}")
    
    column_mapping = {}
    standard_mappings = COLUMN_MAPPINGS[file_type]
    
    for standard_name, variations in standard_mappings.items():
        found_column = None
        
        # First, try exact matches from the predefined variations
        for variation in variations:
            if variation in available_columns:
                found_column = variation
                break
        
        # If no exact match and not in strict mode, try fuzzy matching
        if not found_column and not strict_mode:
            found_column = find_best_column_match(standard_name, available_columns)
        
        if found_column:
            column_mapping[found_column] = standard_name
            logging.info(f"Mapped '{found_column}' -> '{standard_name}'")
    
    return column_mapping


def validate_required_columns(file_type: str, column_mapping: Dict[str, str], 
                             required_columns: Optional[List[str]] = None) -> Tuple[bool, List[str]]:
    """
    Validate that all required columns are present in the mapping.
    
    Args:
        file_type (str): Type of file being validated
        column_mapping (Dict[str, str]): Column mapping from available to standard names
        required_columns (Optional[List[str]]): Override list of required columns
        
    Returns:
        Tuple[bool, List[str]]: (is_valid, list_of_missing_columns)
    """
    # Define required columns for each file type
    default_required = {
        'gl_inquiry': ['Account', 'Job Number', 'Debit', 'Credit'],
        'wip_worksheet': ['Job Number', 'Status'],
        'wip_report': ['Job Number']
    }
    
    if required_columns is None:
        required_columns = default_required.get(file_type, [])
    
    # Check which required columns are missing
    mapped_standard_names = set(column_mapping.values())
    missing_columns = [col for col in required_columns if col not in mapped_standard_names]
    
    is_valid = len(missing_columns) == 0
    
    if not is_valid:
        logging.warning(f"Missing required columns for {file_type}: {missing_columns}")
    else:
        logging.info(f"All required columns found for {file_type}")
    
    return is_valid, missing_columns


def suggest_column_mappings(file_type: str, available_columns: List[str]) -> Dict[str, List[str]]:
    """
    Suggest possible column mappings for manual review.
    
    Args:
        file_type (str): Type of file
        available_columns (List[str]): Available column names
        
    Returns:
        Dict[str, List[str]]: Suggestions for each standard column
    """
    if file_type not in COLUMN_MAPPINGS:
        return {}
    
    suggestions = {}
    standard_mappings = COLUMN_MAPPINGS[file_type]
    
    for standard_name in standard_mappings.keys():
        # Find all columns with similarity > 0.3
        candidates = []
        for available_col in available_columns:
            score = get_similarity_score(standard_name, available_col)
            if score > 0.3:
                candidates.append((available_col, score))
        
        # Sort by score and take top 3
        candidates.sort(key=lambda x: x[1], reverse=True)
        suggestions[standard_name] = [col for col, score in candidates[:3]]
    
    return suggestions


def apply_column_mapping(df, column_mapping: Dict[str, str]):
    """
    Apply column mapping to a DataFrame by renaming columns.
    
    Args:
        df: pandas DataFrame
        column_mapping (Dict[str, str]): Mapping from current to new column names
        
    Returns:
        pandas DataFrame: DataFrame with renamed columns
    """
    # Only rename columns that exist in the DataFrame
    valid_mapping = {old_name: new_name for old_name, new_name in column_mapping.items() 
                    if old_name in df.columns}
    
    if valid_mapping:
        df_renamed = df.rename(columns=valid_mapping)
        logging.info(f"Applied column mapping: {valid_mapping}")
        return df_renamed
    else:
        logging.warning("No valid column mappings found to apply")
        return df.copy()


def get_unmapped_columns(available_columns: List[str], column_mapping: Dict[str, str]) -> List[str]:
    """
    Get list of columns that were not mapped to any standard name.
    
    Args:
        available_columns (List[str]): All available column names
        column_mapping (Dict[str, str]): Applied column mapping
        
    Returns:
        List[str]: List of unmapped column names
    """
    mapped_columns = set(column_mapping.keys())
    unmapped = [col for col in available_columns if col not in mapped_columns]
    
    if unmapped:
        logging.info(f"Unmapped columns: {unmapped}")
    
    return unmapped


def map_dataframe_columns(df, file_type: str, strict_mode: bool = False):
    """
    Convenience function to map DataFrame columns to standard names.
    
    Args:
        df: pandas DataFrame
        file_type (str): Type of file ('gl_inquiry', 'wip_worksheet', 'wip_report')
        strict_mode (bool): If True, only use exact matches
        
    Returns:
        pandas DataFrame: DataFrame with mapped column names
    """
    # Get column mapping
    column_mapping = map_columns_for_file_type(list(df.columns), file_type, strict_mode)
    
    # Apply mapping to DataFrame
    mapped_df = apply_column_mapping(df, column_mapping)
    
    # Validate required columns
    is_valid, missing = validate_required_columns(file_type, column_mapping)
    if not is_valid:
        logging.warning(f"Missing required columns for {file_type}: {missing}")
    
    return mapped_df


if __name__ == "__main__":
    # Example usage and testing
    logging.basicConfig(level=logging.INFO)
    
    # Example: Map GL Inquiry columns
    sample_gl_columns = ['GL Account', 'Job No', 'DR', 'CR', 'Description']
    mapping = map_columns_for_file_type(sample_gl_columns, 'gl_inquiry')
    print("GL Inquiry Mapping:", mapping)
    
    is_valid, missing = validate_required_columns('gl_inquiry', mapping)
    print(f"Valid: {is_valid}, Missing: {missing}")
    
    # Example: Get suggestions
    suggestions = suggest_column_mappings('wip_worksheet', ['Project Number', 'Job Status', 'Project Name'])
    print("Suggestions:", suggestions) 