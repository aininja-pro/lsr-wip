"""
Test cases for Column Mapping Utility Module

This module contains pytest test cases to validate the column mapping functionality.
"""

import pytest
import pandas as pd
from src.data_processing.column_mapping import (
    COLUMN_MAPPINGS,
    get_similarity_score,
    find_best_column_match,
    map_columns_for_file_type,
    validate_required_columns,
    suggest_column_mappings,
    apply_column_mapping,
    get_unmapped_columns
)


class TestGetSimilarityScore:
    """Test cases for get_similarity_score function."""
    
    def test_identical_strings(self):
        """Test similarity score for identical strings."""
        score = get_similarity_score("Job Number", "Job Number")
        assert score == 1.0
    
    def test_case_insensitive(self):
        """Test that comparison is case insensitive."""
        score = get_similarity_score("Job Number", "job number")
        assert score == 1.0
    
    def test_partial_match(self):
        """Test similarity score for partial matches."""
        score = get_similarity_score("Job Number", "Job No")
        assert 0.5 < score < 1.0
    
    def test_no_similarity(self):
        """Test similarity score for completely different strings."""
        score = get_similarity_score("Job Number", "Description")
        assert score < 0.3


class TestFindBestColumnMatch:
    """Test cases for find_best_column_match function."""
    
    def test_exact_match(self):
        """Test finding exact match."""
        available_columns = ['Job Number', 'Status', 'Description']
        match = find_best_column_match('Job Number', available_columns)
        assert match == 'Job Number'
    
    def test_close_match(self):
        """Test finding close match."""
        available_columns = ['Job No', 'Status', 'Description']
        match = find_best_column_match('Job Number', available_columns)
        assert match == 'Job No'
    
    def test_no_good_match(self):
        """Test when no good match exists."""
        available_columns = ['Status', 'Description', 'Amount']
        match = find_best_column_match('Job Number', available_columns, threshold=0.8)
        assert match is None
    
    def test_custom_threshold(self):
        """Test with custom similarity threshold."""
        available_columns = ['Job', 'Status', 'Description']
        
        # With low threshold, should find match
        match_low = find_best_column_match('Job Number', available_columns, threshold=0.3)
        assert match_low == 'Job'
        
        # With high threshold, should not find match
        match_high = find_best_column_match('Job Number', available_columns, threshold=0.8)
        assert match_high is None


class TestMapColumnsForFileType:
    """Test cases for map_columns_for_file_type function."""
    
    def test_gl_inquiry_exact_matches(self):
        """Test mapping GL Inquiry columns with exact matches."""
        available_columns = ['Account', 'Job Number', 'Debit', 'Credit', 'Description']
        mapping = map_columns_for_file_type('gl_inquiry', available_columns)
        
        expected_mapping = {
            'Account': 'Account',
            'Job Number': 'Job Number',
            'Debit': 'Debit',
            'Credit': 'Credit',
            'Description': 'Description'
        }
        assert mapping == expected_mapping
    
    def test_gl_inquiry_variations(self):
        """Test mapping GL Inquiry columns with variations."""
        available_columns = ['GL Account', 'Job No', 'DR', 'CR']
        mapping = map_columns_for_file_type('gl_inquiry', available_columns)
        
        expected_mapping = {
            'GL Account': 'Account',
            'Job No': 'Job Number',
            'DR': 'Debit',
            'CR': 'Credit'
        }
        assert mapping == expected_mapping
    
    def test_wip_worksheet_mapping(self):
        """Test mapping WIP Worksheet columns."""
        available_columns = ['Project Number', 'Job Status', 'Project Name', 'Mat Budget']
        mapping = map_columns_for_file_type('wip_worksheet', available_columns)
        
        expected_mapping = {
            'Project Number': 'Job Number',
            'Job Status': 'Status',
            'Project Name': 'Job Name',
            'Mat Budget': 'Budget Material'
        }
        assert mapping == expected_mapping
    
    def test_strict_mode(self):
        """Test mapping in strict mode (no fuzzy matching)."""
        available_columns = ['JobNum', 'Status']  # 'JobNum' is not in exact variations
        
        # In non-strict mode, should find fuzzy match
        mapping_fuzzy = map_columns_for_file_type('wip_worksheet', available_columns, strict_mode=False)
        assert 'JobNum' in mapping_fuzzy
        
        # In strict mode, should not find fuzzy match
        mapping_strict = map_columns_for_file_type('wip_worksheet', available_columns, strict_mode=True)
        assert 'JobNum' not in mapping_strict
    
    def test_unknown_file_type(self):
        """Test error handling for unknown file type."""
        with pytest.raises(ValueError, match="Unknown file type"):
            map_columns_for_file_type('unknown_type', ['Column1'])
    
    def test_partial_mapping(self):
        """Test mapping when only some columns are available."""
        available_columns = ['Account', 'Job Number']  # Missing Debit and Credit
        mapping = map_columns_for_file_type('gl_inquiry', available_columns)
        
        expected_mapping = {
            'Account': 'Account',
            'Job Number': 'Job Number'
        }
        assert mapping == expected_mapping


class TestValidateRequiredColumns:
    """Test cases for validate_required_columns function."""
    
    def test_all_required_present(self):
        """Test validation when all required columns are present."""
        column_mapping = {
            'Account': 'Account',
            'Job Number': 'Job Number',
            'Debit': 'Debit',
            'Credit': 'Credit'
        }
        
        is_valid, missing = validate_required_columns('gl_inquiry', column_mapping)
        assert is_valid is True
        assert missing == []
    
    def test_missing_required_columns(self):
        """Test validation when required columns are missing."""
        column_mapping = {
            'Account': 'Account',
            'Job Number': 'Job Number'
            # Missing Debit and Credit
        }
        
        is_valid, missing = validate_required_columns('gl_inquiry', column_mapping)
        assert is_valid is False
        assert 'Debit' in missing
        assert 'Credit' in missing
    
    def test_custom_required_columns(self):
        """Test validation with custom required columns."""
        column_mapping = {
            'Job Number': 'Job Number',
            'Status': 'Status'
        }
        
        custom_required = ['Job Number', 'Status', 'Job Name']
        is_valid, missing = validate_required_columns('wip_worksheet', column_mapping, custom_required)
        assert is_valid is False
        assert missing == ['Job Name']
    
    def test_wip_worksheet_validation(self):
        """Test validation for WIP worksheet file type."""
        column_mapping = {
            'Job Number': 'Job Number',
            'Status': 'Status'
        }
        
        is_valid, missing = validate_required_columns('wip_worksheet', column_mapping)
        assert is_valid is True
        assert missing == []


class TestSuggestColumnMappings:
    """Test cases for suggest_column_mappings function."""
    
    def test_suggest_mappings(self):
        """Test column mapping suggestions."""
        available_columns = ['Project Number', 'Job Status', 'Project Name', 'Material Budget']
        suggestions = suggest_column_mappings('wip_worksheet', available_columns)
        
        # Should suggest Project Number for Job Number
        assert 'Job Number' in suggestions
        assert 'Project Number' in suggestions['Job Number']
        
        # Should suggest Job Status for Status
        assert 'Status' in suggestions
        assert 'Job Status' in suggestions['Status']
    
    def test_suggest_mappings_unknown_file_type(self):
        """Test suggestions for unknown file type."""
        suggestions = suggest_column_mappings('unknown_type', ['Column1'])
        assert suggestions == {}
    
    def test_suggest_mappings_no_matches(self):
        """Test suggestions when no good matches exist."""
        available_columns = ['XYZ123', 'ABC456', 'DEF789']  # Very different names
        suggestions = suggest_column_mappings('gl_inquiry', available_columns)
        
        # Should return empty or very few suggestions for each standard column
        for standard_col in COLUMN_MAPPINGS['gl_inquiry'].keys():
            assert standard_col in suggestions
            # Most should be empty, but allow for some very low similarity matches
            assert len(suggestions[standard_col]) <= 3


class TestApplyColumnMapping:
    """Test cases for apply_column_mapping function."""
    
    def test_apply_mapping_success(self):
        """Test successful application of column mapping."""
        df = pd.DataFrame({
            'GL Account': [1, 2, 3],
            'Job No': ['A', 'B', 'C'],
            'DR': [100, 200, 300]
        })
        
        column_mapping = {
            'GL Account': 'Account',
            'Job No': 'Job Number',
            'DR': 'Debit'
        }
        
        result_df = apply_column_mapping(df, column_mapping)
        
        assert 'Account' in result_df.columns
        assert 'Job Number' in result_df.columns
        assert 'Debit' in result_df.columns
        assert 'GL Account' not in result_df.columns
    
    def test_apply_mapping_partial(self):
        """Test application when only some columns exist."""
        df = pd.DataFrame({
            'GL Account': [1, 2, 3],
            'Other Column': ['X', 'Y', 'Z']
        })
        
        column_mapping = {
            'GL Account': 'Account',
            'Nonexistent Column': 'Something'  # This column doesn't exist
        }
        
        result_df = apply_column_mapping(df, column_mapping)
        
        assert 'Account' in result_df.columns
        assert 'Other Column' in result_df.columns
        assert 'Something' not in result_df.columns
    
    def test_apply_mapping_no_valid_mappings(self):
        """Test application when no valid mappings exist."""
        df = pd.DataFrame({
            'Column1': [1, 2, 3],
            'Column2': ['A', 'B', 'C']
        })
        
        column_mapping = {
            'Nonexistent1': 'New1',
            'Nonexistent2': 'New2'
        }
        
        result_df = apply_column_mapping(df, column_mapping)
        
        # Should return copy of original DataFrame
        assert list(result_df.columns) == list(df.columns)
        assert result_df.equals(df)


class TestGetUnmappedColumns:
    """Test cases for get_unmapped_columns function."""
    
    def test_get_unmapped_columns(self):
        """Test identification of unmapped columns."""
        available_columns = ['Account', 'Job Number', 'Description', 'Amount', 'Date']
        column_mapping = {
            'Account': 'Account',
            'Job Number': 'Job Number',
            'Description': 'Description'
        }
        
        unmapped = get_unmapped_columns(available_columns, column_mapping)
        
        assert 'Amount' in unmapped
        assert 'Date' in unmapped
        assert len(unmapped) == 2
    
    def test_get_unmapped_columns_all_mapped(self):
        """Test when all columns are mapped."""
        available_columns = ['Account', 'Job Number']
        column_mapping = {
            'Account': 'Account',
            'Job Number': 'Job Number'
        }
        
        unmapped = get_unmapped_columns(available_columns, column_mapping)
        assert unmapped == []
    
    def test_get_unmapped_columns_none_mapped(self):
        """Test when no columns are mapped."""
        available_columns = ['Column1', 'Column2', 'Column3']
        column_mapping = {}
        
        unmapped = get_unmapped_columns(available_columns, column_mapping)
        assert unmapped == available_columns


class TestColumnMappingsConstant:
    """Test cases for the COLUMN_MAPPINGS constant."""
    
    def test_column_mappings_structure(self):
        """Test that COLUMN_MAPPINGS has expected structure."""
        assert 'gl_inquiry' in COLUMN_MAPPINGS
        assert 'wip_worksheet' in COLUMN_MAPPINGS
        assert 'wip_report' in COLUMN_MAPPINGS
        
        # Each file type should have dictionary of lists
        for file_type, mappings in COLUMN_MAPPINGS.items():
            assert isinstance(mappings, dict)
            for standard_name, variations in mappings.items():
                assert isinstance(variations, list)
                assert len(variations) > 0
    
    def test_required_columns_present(self):
        """Test that required columns are defined in mappings."""
        # GL Inquiry should have required columns
        gl_mappings = COLUMN_MAPPINGS['gl_inquiry']
        assert 'Account' in gl_mappings
        assert 'Job Number' in gl_mappings
        assert 'Debit' in gl_mappings
        assert 'Credit' in gl_mappings
        
        # WIP Worksheet should have required columns
        wip_mappings = COLUMN_MAPPINGS['wip_worksheet']
        assert 'Job Number' in wip_mappings
        assert 'Status' in wip_mappings


if __name__ == "__main__":
    # Run tests if executed directly
    pytest.main([__file__]) 