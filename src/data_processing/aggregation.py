"""
GL Data Aggregation Module

This module handles the aggregation of General Ledger (GL) data from the GL Inquiry Excel file.
It filters accounts containing specific substrings ('5040', '5030', '4020') and computes
aggregated amounts by job number and account type.
"""

import pandas as pd
import logging
from typing import Dict, List, Optional


def load_gl_inquiry(file_path: str) -> pd.DataFrame:
    """
    Load the GL Inquiry Excel file into a pandas DataFrame.
    
    Args:
        file_path (str): Path to the GL Inquiry Excel file
        
    Returns:
        pd.DataFrame: Loaded GL data
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        ValueError: If required columns are missing
    """
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)
        
        # Check for required columns (with common variations)
        required_columns = ['Account', 'Job Number', 'Debit', 'Credit']
        column_variations = {
            'Account': ['Account', 'Account Number', 'Acct', 'GL Account'],
            'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number'],
            'Debit': ['Debit', 'Debit Amount', 'DR', 'Dr'],
            'Credit': ['Credit', 'Credit Amount', 'CR', 'Cr']
        }
        
        # Map column names to standard names
        column_mapping = {}
        for standard_name, variations in column_variations.items():
            found_column = None
            for variation in variations:
                if variation in df.columns:
                    found_column = variation
                    break
            
            if found_column:
                column_mapping[found_column] = standard_name
            else:
                raise ValueError(f"Required column '{standard_name}' not found. Available columns: {list(df.columns)}")
        
        # Rename columns to standard names
        df = df.rename(columns=column_mapping)
        
        logging.info(f"Successfully loaded GL Inquiry file with {len(df)} records")
        return df
        
    except Exception as e:
        logging.error(f"Error loading GL Inquiry file: {str(e)}")
        raise


def filter_gl_accounts(df: pd.DataFrame, account_filters: List[str] = None) -> pd.DataFrame:
    """
    Filter GL data for accounts containing specific substrings.
    
    Args:
        df (pd.DataFrame): GL data DataFrame
        account_filters (List[str]): List of account substrings to filter for
                                   Defaults to ['5040', '5030', '4020']
        
    Returns:
        pd.DataFrame: Filtered GL data
    """
    if account_filters is None:
        account_filters = ['5040', '5030', '4020']
    
    # Convert Account column to string to handle mixed types
    df['Account'] = df['Account'].astype(str)
    
    # Create a boolean mask for accounts containing any of the filter strings
    mask = df['Account'].str.contains('|'.join(account_filters), na=False)
    
    filtered_df = df[mask].copy()
    
    logging.info(f"Filtered GL data from {len(df)} to {len(filtered_df)} records based on account filters: {account_filters}")
    return filtered_df


def compute_amounts(df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute the Amount field as Debit + Credit for each record.
    
    Args:
        df (pd.DataFrame): GL data with Debit and Credit columns
        
    Returns:
        pd.DataFrame: GL data with computed Amount column
    """
    # Fill NaN values with 0 for numeric calculations
    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce').fillna(0)
    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce').fillna(0)
    
    # Compute Amount = Debit + Credit
    df['Amount'] = df['Debit'] + df['Credit']
    
    logging.info("Computed Amount field as Debit + Credit")
    return df


def determine_account_type(account: str) -> str:
    """
    Determine the account type based on the account number.
    
    Args:
        account (str): Account number or code
        
    Returns:
        str: Account type ('Material', 'Sub Labor', 'Other')
    """
    account_str = str(account)
    
    if '5040' in account_str:
        return 'Sub Labor'  # 5040 = Sub Labor costs
    elif '5030' in account_str:
        return 'Material'   # 5030 = Material costs
    elif '4020' in account_str:
        return 'Other'
    else:
        return 'Unknown'


def aggregate_gl_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate GL data by trimmed Job Number and Account Type, summing Amount.
    
    Args:
        df (pd.DataFrame): GL data with Debit and Credit columns
        
    Returns:
        pd.DataFrame: Aggregated GL data grouped by Job Number and Account Type
    """
    # First, compute amounts if not already done
    if 'Amount' not in df.columns:
        df = compute_amounts(df)
    
    # Trim whitespace from Job Number
    df['Job Number'] = df['Job Number'].astype(str).str.strip()
    
    # Determine account type for each record
    df['Account Type'] = df['Account'].apply(determine_account_type)
    
    # Group by Job Number and Account Type, sum the Amount
    aggregated_df = df.groupby(['Job Number', 'Account Type'])['Amount'].sum().reset_index()
    
    # Pivot to have Account Types as columns
    pivot_df = aggregated_df.pivot(index='Job Number', columns='Account Type', values='Amount').fillna(0)
    
    # Reset index to make Job Number a regular column
    pivot_df = pivot_df.reset_index()
    
    # Ensure we have the expected columns
    expected_columns = ['Material', 'Sub Labor', 'Other']
    for col in expected_columns:
        if col not in pivot_df.columns:
            pivot_df[col] = 0
    
    logging.info(f"Aggregated GL data to {len(pivot_df)} job records")
    return pivot_df


def process_gl_inquiry(file_path: str, account_filters: List[str] = None) -> pd.DataFrame:
    """
    Complete processing pipeline for GL Inquiry data.
    
    Args:
        file_path (str): Path to the GL Inquiry Excel file
        account_filters (List[str]): Account filter strings (default: ['5040', '5030', '4020'])
        
    Returns:
        pd.DataFrame: Processed and aggregated GL data
    """
    # Load the GL Inquiry file
    gl_data = load_gl_inquiry(file_path)
    
    # Filter for specific accounts
    filtered_data = filter_gl_accounts(gl_data, account_filters)
    
    # Compute amounts
    amounts_data = compute_amounts(filtered_data)
    
    # Aggregate by job and account type
    aggregated_data = aggregate_gl_data(amounts_data)
    
    logging.info("GL Inquiry processing pipeline completed successfully")
    return aggregated_data


if __name__ == "__main__":
    # Example usage and testing
    logging.basicConfig(level=logging.INFO)
    
    # This would be used for testing with a sample file
    # result = process_gl_inquiry('sample_gl_inquiry.xlsx')
    # print(result.head())
    pass 