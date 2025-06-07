"""
Test cases for GL Data Aggregation Module

This module contains pytest test cases to validate the GL aggregation functionality.
"""

import pytest
import pandas as pd
import numpy as np
import tempfile
import os
from src.data_processing.aggregation import (
    load_gl_inquiry,
    filter_gl_accounts,
    compute_amounts,
    determine_account_type,
    aggregate_gl_data,
    process_gl_inquiry
)


@pytest.fixture
def sample_gl_data():
    """Create sample GL data for testing."""
    return pd.DataFrame({
        'Account': ['5040-001', '5030-002', '4020-003', '6000-001', '5040-002'],
        'Job Number': ['  JOB001  ', 'JOB002', 'JOB001', 'JOB003', 'JOB001'],
        'Debit': [1000.00, 500.00, 0.00, 750.00, 250.00],
        'Credit': [0.00, 0.00, 300.00, 0.00, 0.00]
    })


@pytest.fixture
def sample_excel_file():
    """Create a temporary Excel file for testing."""
    data = {
        'Account': ['5040-001', '5030-002', '4020-003', '6000-001', '5040-002'],
        'Job Number': ['JOB001', 'JOB002', 'JOB001', 'JOB003', 'JOB001'],
        'Debit': [1000.00, 500.00, 0.00, 750.00, 250.00],
        'Credit': [0.00, 0.00, 300.00, 0.00, 0.00]
    }
    df = pd.DataFrame(data)
    
    # Create temporary file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
        df.to_excel(tmp_file.name, index=False)
        yield tmp_file.name
    
    # Cleanup
    os.unlink(tmp_file.name)


class TestLoadGLInquiry:
    """Test cases for load_gl_inquiry function."""
    
    def test_load_gl_inquiry_success(self, sample_excel_file):
        """Test successful loading of GL Inquiry file."""
        df = load_gl_inquiry(sample_excel_file)
        
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 5
        assert 'Account' in df.columns
        assert 'Job Number' in df.columns
        assert 'Debit' in df.columns
        assert 'Credit' in df.columns
    
    def test_load_gl_inquiry_column_variations(self):
        """Test loading with different column name variations."""
        data = {
            'GL Account': ['5040-001', '5030-002'],
            'Job No': ['JOB001', 'JOB002'],
            'DR': [1000.00, 500.00],
            'CR': [0.00, 0.00]
        }
        df = pd.DataFrame(data)
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            df.to_excel(tmp_file.name, index=False)
            
            try:
                result_df = load_gl_inquiry(tmp_file.name)
                
                # Check that columns were renamed correctly
                assert 'Account' in result_df.columns
                assert 'Job Number' in result_df.columns
                assert 'Debit' in result_df.columns
                assert 'Credit' in result_df.columns
                
            finally:
                os.unlink(tmp_file.name)
    
    def test_load_gl_inquiry_missing_columns(self):
        """Test error handling when required columns are missing."""
        data = {
            'Account': ['5040-001', '5030-002'],
            'Debit': [1000.00, 500.00]
            # Missing Job Number and Credit columns
        }
        df = pd.DataFrame(data)
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            df.to_excel(tmp_file.name, index=False)
            
            try:
                with pytest.raises(ValueError, match="Required column"):
                    load_gl_inquiry(tmp_file.name)
            finally:
                os.unlink(tmp_file.name)
    
    def test_load_gl_inquiry_file_not_found(self):
        """Test error handling when file doesn't exist."""
        with pytest.raises(FileNotFoundError):
            load_gl_inquiry('nonexistent_file.xlsx')


class TestFilterGLAccounts:
    """Test cases for filter_gl_accounts function."""
    
    def test_filter_gl_accounts_default_filters(self, sample_gl_data):
        """Test filtering with default account filters."""
        filtered_df = filter_gl_accounts(sample_gl_data)
        
        # Should include accounts with 5040, 5030, 4020 but exclude 6000
        assert len(filtered_df) == 4
        assert '6000-001' not in filtered_df['Account'].values
    
    def test_filter_gl_accounts_custom_filters(self, sample_gl_data):
        """Test filtering with custom account filters."""
        custom_filters = ['5040', '6000']
        filtered_df = filter_gl_accounts(sample_gl_data, custom_filters)
        
        # Should only include accounts with 5040 and 6000
        assert len(filtered_df) == 3  # Two 5040 records and one 6000 record
        assert all('5040' in str(acc) or '6000' in str(acc) for acc in filtered_df['Account'])
    
    def test_filter_gl_accounts_no_matches(self, sample_gl_data):
        """Test filtering when no accounts match the filters."""
        custom_filters = ['9999']
        filtered_df = filter_gl_accounts(sample_gl_data, custom_filters)
        
        assert len(filtered_df) == 0


class TestComputeAmounts:
    """Test cases for compute_amounts function."""
    
    def test_compute_amounts_normal_case(self, sample_gl_data):
        """Test normal amount computation."""
        result_df = compute_amounts(sample_gl_data)
        
        assert 'Amount' in result_df.columns
        # First row: 1000 + 0 = 1000
        assert result_df.iloc[0]['Amount'] == 1000.00
        # Third row: 0 + 300 = 300
        assert result_df.iloc[2]['Amount'] == 300.00
    
    def test_compute_amounts_with_nan_values(self):
        """Test amount computation with NaN values."""
        data = pd.DataFrame({
            'Account': ['5040-001', '5030-002'],
            'Job Number': ['JOB001', 'JOB002'],
            'Debit': [1000.00, np.nan],
            'Credit': [np.nan, 500.00]
        })
        
        result_df = compute_amounts(data)
        
        # NaN values should be treated as 0
        assert result_df.iloc[0]['Amount'] == 1000.00  # 1000 + 0
        assert result_df.iloc[1]['Amount'] == 500.00   # 0 + 500
    
    def test_compute_amounts_string_values(self):
        """Test amount computation with string values that can't be converted."""
        data = pd.DataFrame({
            'Account': ['5040-001', '5030-002'],
            'Job Number': ['JOB001', 'JOB002'],
            'Debit': ['1000', 'invalid'],
            'Credit': [0, '500']
        })
        
        result_df = compute_amounts(data)
        
        # Valid strings should be converted, invalid ones should become 0
        assert result_df.iloc[0]['Amount'] == 1000.00  # '1000' + 0
        assert result_df.iloc[1]['Amount'] == 500.00   # 0 + '500'


class TestDetermineAccountType:
    """Test cases for determine_account_type function."""
    
    def test_determine_account_type_material(self):
        """Test identification of Material accounts."""
        assert determine_account_type('5040-001') == 'Material'
        assert determine_account_type('ABC-5040-XYZ') == 'Material'
    
    def test_determine_account_type_labor(self):
        """Test identification of Labor accounts."""
        assert determine_account_type('5030-002') == 'Labor'
        assert determine_account_type('DEF-5030-ABC') == 'Labor'
    
    def test_determine_account_type_other(self):
        """Test identification of Other accounts."""
        assert determine_account_type('4020-003') == 'Other'
        assert determine_account_type('XYZ-4020-DEF') == 'Other'
    
    def test_determine_account_type_unknown(self):
        """Test identification of Unknown account types."""
        assert determine_account_type('6000-001') == 'Unknown'
        assert determine_account_type('1234-567') == 'Unknown'


class TestAggregateGLData:
    """Test cases for aggregate_gl_data function."""
    
    def test_aggregate_gl_data_basic(self):
        """Test basic aggregation functionality."""
        data = pd.DataFrame({
            'Account': ['5040-001', '5030-002', '5040-003', '4020-001'],
            'Job Number': ['  JOB001  ', 'JOB001', 'JOB002', 'JOB001'],
            'Amount': [1000.00, 500.00, 750.00, 300.00]
        })
        
        result_df = aggregate_gl_data(data)
        
        # Check that Job Numbers are trimmed
        assert '  JOB001  ' not in result_df['Job Number'].values
        assert 'JOB001' in result_df['Job Number'].values
        assert 'JOB002' in result_df['Job Number'].values
        
        # Check aggregation results
        job001_row = result_df[result_df['Job Number'] == 'JOB001'].iloc[0]
        assert job001_row['Material'] == 1000.00  # Only one Material entry for JOB001
        assert job001_row['Labor'] == 500.00      # Only one Labor entry for JOB001
        assert job001_row['Other'] == 300.00      # Only one Other entry for JOB001
        
        job002_row = result_df[result_df['Job Number'] == 'JOB002'].iloc[0]
        assert job002_row['Material'] == 750.00   # One Material entry for JOB002
        assert job002_row['Labor'] == 0.00        # No Labor entries for JOB002
        assert job002_row['Other'] == 0.00        # No Other entries for JOB002
    
    def test_aggregate_gl_data_multiple_same_type(self):
        """Test aggregation when multiple records of same type exist for one job."""
        data = pd.DataFrame({
            'Account': ['5040-001', '5040-002', '5030-001'],
            'Job Number': ['JOB001', 'JOB001', 'JOB001'], 
            'Amount': [1000.00, 500.00, 750.00]
        })
        
        result_df = aggregate_gl_data(data)
        
        job001_row = result_df[result_df['Job Number'] == 'JOB001'].iloc[0]
        assert job001_row['Material'] == 1500.00  # 1000 + 500
        assert job001_row['Labor'] == 750.00      # Single Labor entry
        assert job001_row['Other'] == 0.00        # No Other entries


class TestProcessGLInquiry:
    """Test cases for the complete process_gl_inquiry pipeline."""
    
    def test_process_gl_inquiry_complete_pipeline(self, sample_excel_file):
        """Test the complete GL processing pipeline."""
        result_df = process_gl_inquiry(sample_excel_file)
        
        # Check that result is a DataFrame with expected structure
        assert isinstance(result_df, pd.DataFrame)
        assert 'Job Number' in result_df.columns
        assert 'Material' in result_df.columns
        assert 'Labor' in result_df.columns
        assert 'Other' in result_df.columns
        
        # Verify that only filtered accounts are included in aggregation
        # From sample data: 3 records should match filters (5040, 5030, 4020)
        # JOB001 has Material=1250 (1000+250), Other=300
        # JOB002 has Labor=500
        assert len(result_df) == 2  # JOB001 and JOB002
    
    def test_process_gl_inquiry_with_custom_filters(self, sample_excel_file):
        """Test pipeline with custom account filters."""
        custom_filters = ['5040']  # Only Material accounts
        result_df = process_gl_inquiry(sample_excel_file, custom_filters)
        
        # Should only process Material accounts
        # All Labor and Other amounts should be 0
        assert all(result_df['Labor'] == 0)
        assert all(result_df['Other'] == 0)
        assert any(result_df['Material'] > 0)  # Should have some Material amounts


if __name__ == "__main__":
    # Run tests if executed directly
    pytest.main([__file__]) 