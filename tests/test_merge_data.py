"""
Test cases for Data Merging Module

This module contains pytest test cases to validate the WIP worksheet merging functionality.
"""

import pytest
import pandas as pd
import numpy as np
import tempfile
import os
from src.data_processing.merge_data import (
    load_wip_worksheet,
    trim_job_numbers,
    filter_closed_jobs,
    merge_wip_with_gl,
    compute_variances,
    process_wip_merge,
    get_jobs_for_update
)


@pytest.fixture
def sample_wip_data():
    """Create sample WIP worksheet data for testing."""
    return pd.DataFrame({
        'Job Number': ['  JOB001  ', 'JOB002', 'JOB003', 'JOB004'],
        'Status': ['Active', 'Closed', 'Active', 'CLOSED'],
        'Job Name': ['Project Alpha', 'Project Beta', 'Project Gamma', 'Project Delta'],
        'Budget Material': [10000.00, 5000.00, 7500.00, 3000.00],
        'Budget Labor': [8000.00, 4000.00, 6000.00, 2500.00]
    })


@pytest.fixture
def sample_gl_data():
    """Create sample GL aggregated data for testing."""
    return pd.DataFrame({
        'Job Number': ['JOB001', 'JOB002', 'JOB005'],
        'Material': [9500.00, 5200.00, 1000.00],
        'Labor': [8200.00, 3800.00, 500.00],
        'Other': [200.00, 100.00, 50.00]
    })


@pytest.fixture
def sample_wip_excel_file():
    """Create a temporary WIP worksheet Excel file for testing."""
    data = {
        'Job Number': ['JOB001', 'JOB002', 'JOB003'],
        'Status': ['Active', 'Closed', 'Active'],
        'Job Name': ['Project Alpha', 'Project Beta', 'Project Gamma'],
        'Budget Material': [10000.00, 5000.00, 7500.00],
        'Budget Labor': [8000.00, 4000.00, 6000.00]
    }
    df = pd.DataFrame(data)
    
    # Create temporary file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
        df.to_excel(tmp_file.name, index=False)
        yield tmp_file.name
    
    # Cleanup
    os.unlink(tmp_file.name)


class TestLoadWIPWorksheet:
    """Test cases for load_wip_worksheet function."""
    
    def test_load_wip_worksheet_success(self, sample_wip_excel_file):
        """Test successful loading of WIP Worksheet file."""
        df = load_wip_worksheet(sample_wip_excel_file)
        
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 3
        assert 'Job Number' in df.columns
        assert 'Status' in df.columns
        assert 'Job Name' in df.columns
    
    def test_load_wip_worksheet_column_variations(self):
        """Test loading with different column name variations."""
        data = {
            'Job No': ['JOB001', 'JOB002'],
            'Job Status': ['Active', 'Closed'],
            'Project Name': ['Alpha', 'Beta'],
            'Mat Budget': [10000.00, 5000.00],
            'Lab Budget': [8000.00, 4000.00]
        }
        df = pd.DataFrame(data)
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            df.to_excel(tmp_file.name, index=False)
            
            try:
                result_df = load_wip_worksheet(tmp_file.name)
                
                # Check that columns were renamed correctly
                assert 'Job Number' in result_df.columns
                assert 'Status' in result_df.columns
                assert 'Job Name' in result_df.columns
                assert 'Budget Material' in result_df.columns
                assert 'Budget Labor' in result_df.columns
                
            finally:
                os.unlink(tmp_file.name)
    
    def test_load_wip_worksheet_missing_required_columns(self):
        """Test error handling when required columns are missing."""
        data = {
            'Job Number': ['JOB001', 'JOB002'],
            'Job Name': ['Alpha', 'Beta']
            # Missing Status column
        }
        df = pd.DataFrame(data)
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            df.to_excel(tmp_file.name, index=False)
            
            try:
                with pytest.raises(ValueError, match="Required column"):
                    load_wip_worksheet(tmp_file.name)
            finally:
                os.unlink(tmp_file.name)


class TestTrimJobNumbers:
    """Test cases for trim_job_numbers function."""
    
    def test_trim_job_numbers_basic(self, sample_wip_data):
        """Test basic job number trimming."""
        result_df = trim_job_numbers(sample_wip_data)
        
        # Check that whitespace is trimmed
        assert '  JOB001  ' not in result_df['Job Number'].values
        assert 'JOB001' in result_df['Job Number'].values
        assert all(job.strip() == job for job in result_df['Job Number'])
    
    def test_trim_job_numbers_preserves_other_columns(self, sample_wip_data):
        """Test that trimming preserves other columns."""
        original_columns = set(sample_wip_data.columns)
        result_df = trim_job_numbers(sample_wip_data)
        
        assert set(result_df.columns) == original_columns
        assert len(result_df) == len(sample_wip_data)


class TestFilterClosedJobs:
    """Test cases for filter_closed_jobs function."""
    
    def test_filter_closed_jobs_exclude_by_default(self, sample_wip_data):
        """Test that closed jobs are excluded by default."""
        result_df = filter_closed_jobs(sample_wip_data, include_closed=False)
        
        # Should exclude JOB002 and JOB004 (both closed)
        assert len(result_df) == 2
        assert '  JOB001  ' in result_df['Job Number'].values  # Note: whitespace preserved
        assert 'JOB003' in result_df['Job Number'].values
        assert 'JOB002' not in result_df['Job Number'].values
        assert 'JOB004' not in result_df['Job Number'].values
    
    def test_filter_closed_jobs_include_when_requested(self, sample_wip_data):
        """Test that closed jobs are included when requested."""
        result_df = filter_closed_jobs(sample_wip_data, include_closed=True)
        
        # Should include all jobs
        assert len(result_df) == len(sample_wip_data)
        assert 'JOB002' in result_df['Job Number'].values
        assert 'JOB004' in result_df['Job Number'].values
    
    def test_filter_closed_jobs_case_insensitive(self):
        """Test that filtering is case insensitive."""
        data = pd.DataFrame({
            'Job Number': ['JOB001', 'JOB002', 'JOB003'],
            'Status': ['Active', 'CLOSED', 'closed']
        })
        
        result_df = filter_closed_jobs(data, include_closed=False)
        
        # Should exclude both CLOSED and closed
        assert len(result_df) == 1
        assert 'JOB001' in result_df['Job Number'].values


class TestMergeWIPWithGL:
    """Test cases for merge_wip_with_gl function."""
    
    def test_merge_wip_with_gl_basic(self, sample_wip_data, sample_gl_data):
        """Test basic merging functionality."""
        # Trim job numbers first
        wip_trimmed = trim_job_numbers(sample_wip_data)
        result_df = merge_wip_with_gl(wip_trimmed, sample_gl_data)
        
        # Should have all WIP records (left join)
        assert len(result_df) == len(sample_wip_data)
        
        # Check that GL data was merged correctly
        job001_row = result_df[result_df['Job Number'] == 'JOB001'].iloc[0]
        assert job001_row['Material'] == 9500.00
        assert job001_row['Labor'] == 8200.00
        assert job001_row['Other'] == 200.00
        
        # Check that missing GL data is filled with 0
        job003_row = result_df[result_df['Job Number'] == 'JOB003'].iloc[0]
        assert job003_row['Material'] == 0.00  # No GL data for JOB003
        assert job003_row['Labor'] == 0.00
        assert job003_row['Other'] == 0.00
    
    def test_merge_wip_with_gl_no_fill_zeros(self, sample_wip_data, sample_gl_data):
        """Test merging without filling missing values with zeros."""
        wip_trimmed = trim_job_numbers(sample_wip_data)
        result_df = merge_wip_with_gl(wip_trimmed, sample_gl_data, fill_missing_with_zero=False)
        
        # Check that missing GL data is NaN
        job003_row = result_df[result_df['Job Number'] == 'JOB003'].iloc[0]
        assert pd.isna(job003_row['Material'])
        assert pd.isna(job003_row['Labor'])
        assert pd.isna(job003_row['Other'])
    
    def test_merge_wip_with_gl_trimming(self):
        """Test that job numbers are trimmed during merge."""
        wip_data = pd.DataFrame({
            'Job Number': ['  JOB001  ', 'JOB002'],
            'Status': ['Active', 'Active']
        })
        
        gl_data = pd.DataFrame({
            'Job Number': ['JOB001', '  JOB002  '],
            'Material': [1000.00, 2000.00],
            'Labor': [500.00, 1000.00],
            'Other': [100.00, 200.00]
        })
        
        result_df = merge_wip_with_gl(wip_data, gl_data)
        
        # Both jobs should have GL data despite whitespace differences
        assert result_df.iloc[0]['Material'] == 1000.00
        assert result_df.iloc[1]['Material'] == 2000.00


class TestComputeVariances:
    """Test cases for compute_variances function."""
    
    def test_compute_variances_basic(self):
        """Test basic variance computation."""
        data = pd.DataFrame({
            'Job Number': ['JOB001', 'JOB002'],
            'Material': [9500.00, 5200.00],
            'Labor': [8200.00, 3800.00],
            'Budget Material': [10000.00, 5000.00],
            'Budget Labor': [8000.00, 4000.00]
        })
        
        result_df = compute_variances(data)
        
        # Check Material variance (Actual - Budget)
        assert result_df.iloc[0]['Material Variance'] == -500.00  # 9500 - 10000
        assert result_df.iloc[1]['Material Variance'] == 200.00   # 5200 - 5000
        
        # Check Labor variance (Actual - Budget)
        assert result_df.iloc[0]['Labor Variance'] == 200.00      # 8200 - 8000
        assert result_df.iloc[1]['Labor Variance'] == -200.00     # 3800 - 4000
        
        # Check Total variance
        assert result_df.iloc[0]['Total Variance'] == -300.00     # -500 + 200
        assert result_df.iloc[1]['Total Variance'] == 0.00        # 200 + (-200)
    
    def test_compute_variances_missing_budget_columns(self):
        """Test variance computation when budget columns are missing."""
        data = pd.DataFrame({
            'Job Number': ['JOB001', 'JOB002'],
            'Material': [9500.00, 5200.00],
            'Labor': [8200.00, 3800.00]
            # No budget columns
        })
        
        result_df = compute_variances(data)
        
        # Should not add variance columns if budget columns are missing
        assert 'Material Variance' not in result_df.columns
        assert 'Labor Variance' not in result_df.columns
        assert 'Total Variance' not in result_df.columns
    
    def test_compute_variances_with_nan_budget(self):
        """Test variance computation with NaN budget values."""
        data = pd.DataFrame({
            'Job Number': ['JOB001', 'JOB002'],
            'Material': [9500.00, 5200.00],
            'Budget Material': [10000.00, np.nan]
        })
        
        result_df = compute_variances(data)
        
        # NaN budget should be treated as 0
        assert result_df.iloc[0]['Material Variance'] == -500.00  # 9500 - 10000
        assert result_df.iloc[1]['Material Variance'] == 5200.00  # 5200 - 0


class TestGetJobsForUpdate:
    """Test cases for get_jobs_for_update function."""
    
    def test_get_jobs_for_update_material_section(self):
        """Test getting jobs for Material section (5040)."""
        data = pd.DataFrame({
            'Job Number': ['JOB001', 'JOB002'],
            'Job Name': ['Alpha', 'Beta'],
            'Material': [9500.00, 5200.00],
            'Labor': [8200.00, 3800.00],
            'Budget Material': [10000.00, 5000.00],
            'Material Variance': [-500.00, 200.00]
        })
        
        result_df = get_jobs_for_update(data, '5040')
        
        # Should include Material-related columns
        expected_columns = ['Job Number', 'Job Name', 'Material', 'Budget Material', 'Material Variance']
        assert all(col in result_df.columns for col in expected_columns)
        assert 'Labor' not in result_df.columns
    
    def test_get_jobs_for_update_labor_section(self):
        """Test getting jobs for Labor section (5030)."""
        data = pd.DataFrame({
            'Job Number': ['JOB001', 'JOB002'],
            'Job Name': ['Alpha', 'Beta'],
            'Material': [9500.00, 5200.00],
            'Labor': [8200.00, 3800.00],
            'Budget Labor': [8000.00, 4000.00],
            'Labor Variance': [200.00, -200.00]
        })
        
        result_df = get_jobs_for_update(data, '5030')
        
        # Should include Labor-related columns
        expected_columns = ['Job Number', 'Job Name', 'Labor', 'Budget Labor', 'Labor Variance']
        assert all(col in result_df.columns for col in expected_columns)
        assert 'Material' not in result_df.columns
    
    def test_get_jobs_for_update_invalid_section(self):
        """Test error handling for invalid section type."""
        data = pd.DataFrame({'Job Number': ['JOB001']})
        
        with pytest.raises(ValueError, match="Unknown section type"):
            get_jobs_for_update(data, '9999')


class TestProcessWIPMerge:
    """Test cases for the complete process_wip_merge pipeline."""
    
    def test_process_wip_merge_complete_pipeline(self, sample_wip_excel_file, sample_gl_data):
        """Test the complete WIP merge processing pipeline."""
        result_df = process_wip_merge(sample_wip_excel_file, sample_gl_data, include_closed=False)
        
        # Check that result is a DataFrame with expected structure
        assert isinstance(result_df, pd.DataFrame)
        assert 'Job Number' in result_df.columns
        assert 'Material' in result_df.columns
        assert 'Labor' in result_df.columns
        
        # Should exclude closed jobs by default
        assert 'JOB002' not in result_df['Job Number'].values  # JOB002 is closed
        assert 'JOB001' in result_df['Job Number'].values      # JOB001 is active
        assert 'JOB003' in result_df['Job Number'].values      # JOB003 is active
    
    def test_process_wip_merge_include_closed(self, sample_wip_excel_file, sample_gl_data):
        """Test pipeline with closed jobs included."""
        result_df = process_wip_merge(sample_wip_excel_file, sample_gl_data, include_closed=True)
        
        # Should include all jobs including closed ones
        assert len(result_df) == 3  # All jobs from sample file
        assert 'JOB002' in result_df['Job Number'].values  # Closed job included


if __name__ == "__main__":
    # Run tests if executed directly
    pytest.main([__file__]) 