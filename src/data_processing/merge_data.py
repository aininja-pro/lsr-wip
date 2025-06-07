"""
Data Merging Module

This module handles loading the WIP Worksheet Excel file, trimming job numbers,
and merging with aggregated GL data. It also handles filtering of closed jobs.
"""

import pandas as pd
import logging
from typing import Optional, Dict, List


def load_wip_worksheet(file_path: str) -> pd.DataFrame:
    """
    Load the WIP Worksheet Excel file into a pandas DataFrame.
    
    Args:
        file_path (str): Path to the WIP Worksheet Excel file
        
    Returns:
        pd.DataFrame: Loaded WIP data
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        ValueError: If required columns are missing
    """
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)
        
        # Check for required columns (with common variations)
        column_variations = {
            'Job Number': ['Job Number', 'Job No', 'Job #', 'Job', 'Project Number', 'Project No'],
            'Status': ['Status', 'Job Status', 'Project Status', 'State'],
            'Job Name': ['Job Name', 'Project Name', 'Description', 'Job Description'],
            'Budget Material': ['Budget Material', 'Material Budget', 'Mat Budget', 'Budget Mat'],
            'Budget Labor': ['Budget Labor', 'Labor Budget', 'Lab Budget', 'Budget Lab'],
            'Actual Material': ['Actual Material', 'Material Actual', 'Mat Actual', 'Actual Mat'],
            'Actual Labor': ['Actual Labor', 'Labor Actual', 'Lab Actual', 'Actual Lab']
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
                # Some columns might be optional, only require Job Number and Status
                if standard_name in ['Job Number', 'Status']:
                    raise ValueError(f"Required column '{standard_name}' not found. Available columns: {list(df.columns)}")
        
        # Rename columns to standard names
        df = df.rename(columns=column_mapping)
        
        logging.info(f"Successfully loaded WIP Worksheet file with {len(df)} records")
        return df
        
    except Exception as e:
        logging.error(f"Error loading WIP Worksheet file: {str(e)}")
        raise


def trim_job_numbers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Trim whitespace from Job Number column.
    
    Args:
        df (pd.DataFrame): WIP data DataFrame
        
    Returns:
        pd.DataFrame: WIP data with trimmed job numbers
    """
    df = df.copy()
    df['Job Number'] = df['Job Number'].astype(str).str.strip()
    
    logging.info("Trimmed whitespace from Job Number column")
    return df


def filter_closed_jobs(df: pd.DataFrame, include_closed: bool = False) -> pd.DataFrame:
    """
    Filter out closed jobs unless specifically requested to include them.
    
    Args:
        df (pd.DataFrame): WIP data DataFrame
        include_closed (bool): Whether to include closed jobs (default: False)
        
    Returns:
        pd.DataFrame: Filtered WIP data
    """
    if include_closed:
        logging.info("Including all jobs (including closed)")
        return df
    
    # Filter out closed jobs (case insensitive)
    df = df.copy()
    df['Status'] = df['Status'].astype(str).str.lower()
    filtered_df = df[df['Status'] != 'closed'].copy()
    
    closed_count = len(df) - len(filtered_df)
    logging.info(f"Filtered out {closed_count} closed jobs. Remaining: {len(filtered_df)} jobs")
    
    return filtered_df


def merge_wip_with_gl(wip_df: pd.DataFrame, gl_df: pd.DataFrame, 
                      fill_missing_with_zero: bool = True) -> pd.DataFrame:
    """
    Left-join WIP Worksheet data with aggregated GL data by Job Number.
    
    Args:
        wip_df (pd.DataFrame): WIP Worksheet data
        gl_df (pd.DataFrame): Aggregated GL data
        fill_missing_with_zero (bool): Whether to fill missing GL values with 0
        
    Returns:
        pd.DataFrame: Merged WIP and GL data
    """
    # Ensure both dataframes have trimmed job numbers
    wip_df = wip_df.copy()
    gl_df = gl_df.copy()
    
    wip_df['Job Number'] = wip_df['Job Number'].astype(str).str.strip()
    gl_df['Job Number'] = gl_df['Job Number'].astype(str).str.strip()
    
    # Perform left join
    merged_df = pd.merge(wip_df, gl_df, on='Job Number', how='left')
    
    # Fill missing GL values with 0 if requested
    if fill_missing_with_zero:
        gl_columns = ['Material', 'Sub Labor', 'Other']
        for col in gl_columns:
            if col in merged_df.columns:
                merged_df[col] = merged_df[col].fillna(0)
    
    logging.info(f"Merged WIP data ({len(wip_df)} records) with GL data ({len(gl_df)} records). Result: {len(merged_df)} records")
    
    return merged_df


def compute_variances(df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute variances between Actual and Budget amounts for Material and Labor.
    
    Args:
        df (pd.DataFrame): Merged WIP and GL data
        
    Returns:
        pd.DataFrame: Data with computed variance columns
    """
    df = df.copy()
    
    # Compute Material variance (Actual - Budget)
    if 'Material' in df.columns and 'Budget Material' in df.columns:
        df['Budget Material'] = pd.to_numeric(df['Budget Material'], errors='coerce').fillna(0)
        df['Material Variance'] = df['Material'] - df['Budget Material']
    
    # Compute Sub Labor variance (Actual - Budget)
    if 'Sub Labor' in df.columns and 'Budget Labor' in df.columns:
        df['Budget Labor'] = pd.to_numeric(df['Budget Labor'], errors='coerce').fillna(0)
        df['Sub Labor Variance'] = df['Sub Labor'] - df['Budget Labor']
    
    # Compute total variance
    variance_columns = [col for col in df.columns if 'Variance' in col]
    if variance_columns:
        df['Total Variance'] = df[variance_columns].sum(axis=1)
    
    logging.info("Computed variance columns")
    return df


def process_wip_merge(wip_file_path: str, gl_df: pd.DataFrame, 
                      include_closed: bool = False,
                      fill_missing_with_zero: bool = True) -> pd.DataFrame:
    """
    Complete processing pipeline for merging WIP Worksheet with GL data.
    
    Args:
        wip_file_path (str): Path to the WIP Worksheet Excel file
        gl_df (pd.DataFrame): Aggregated GL data
        include_closed (bool): Whether to include closed jobs
        fill_missing_with_zero (bool): Whether to fill missing GL values with 0
        
    Returns:
        pd.DataFrame: Processed and merged data
    """
    # Load WIP Worksheet
    wip_data = load_wip_worksheet(wip_file_path)
    
    # Trim job numbers
    wip_data = trim_job_numbers(wip_data)
    
    # Filter closed jobs if requested
    if not include_closed:
        wip_data = filter_closed_jobs(wip_data, include_closed=False)
    
    # Merge with GL data
    merged_data = merge_wip_with_gl(wip_data, gl_df, fill_missing_with_zero)
    
    # Compute variances
    final_data = compute_variances(merged_data)
    
    logging.info("WIP merge processing pipeline completed successfully")
    return final_data


def get_jobs_for_update(merged_df: pd.DataFrame, section_type: str) -> pd.DataFrame:
    """
    Get jobs that need to be updated in a specific section of the WIP Report.
    
    Args:
        merged_df (pd.DataFrame): Merged WIP and GL data
        section_type (str): Type of section ('5040' for Material, '5030' for Labor)
        
    Returns:
        pd.DataFrame: Jobs data for the specified section
    """
    if section_type == '5040':
        # Sub Labor section - return jobs with Sub Labor column
        relevant_columns = ['Job Number', 'Job Name', 'Sub Labor', 'Budget Labor', 'Sub Labor Variance']
        available_columns = [col for col in relevant_columns if col in merged_df.columns]
        return merged_df[available_columns].copy()
    
    elif section_type == '5030':
        # Material section - return jobs with Material column
        relevant_columns = ['Job Number', 'Job Name', 'Material', 'Budget Material', 'Material Variance']
        available_columns = [col for col in relevant_columns if col in merged_df.columns]
        return merged_df[available_columns].copy()
    
    else:
        raise ValueError(f"Unknown section type: {section_type}. Expected '5040' or '5030'")


if __name__ == "__main__":
    # Example usage and testing
    logging.basicConfig(level=logging.INFO)
    
    # This would be used for testing with sample files
    # Sample GL data
    # gl_data = pd.DataFrame({
    #     'Job Number': ['JOB001', 'JOB002'],
    #     'Material': [1000, 500],
    #     'Labor': [750, 300],
    #     'Other': [200, 100]
    # })
    # 
    # result = process_wip_merge('sample_wip_worksheet.xlsx', gl_data)
    # print(result.head())
    pass 