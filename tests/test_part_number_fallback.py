"""
Tests for the part-number free fallback chain:
  1. extract_part_number_from_filename  (filename_utils.py)
  2. PL Part Number + Filename fallback integration
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.services.extraction.filename_utils import extract_part_number_from_filename


# =====================================================================
# extract_part_number_from_filename  –  unit tests
# =====================================================================

class TestExtractPartNumberFromFilename:
    """Test the regex-based filename part-number extraction."""

    # --- Positive cases ---

    def test_standard_iai_filename(self):
        """33956_MA-S9160-1000_SHEET1_04112025.pdf → MA-S9160-1000"""
        result = extract_part_number_from_filename(
            r"C:\drawings\33956_MA-S9160-1000_SHEET1_04112025.pdf"
        )
        assert result == "MA-S9160-1000"

    def test_simple_part_number_only(self):
        """MA-S9160-1000.pdf → MA-S9160-1000"""
        result = extract_part_number_from_filename("MA-S9160-1000.pdf")
        assert result == "MA-S9160-1000"

    def test_dwg_prefix(self):
        """DWG_12345-678-901_REV_A.pdf → 12345-678-901"""
        result = extract_part_number_from_filename("DWG_12345-678-901_REV_A.pdf")
        assert result == "12345-678-901"

    def test_two_candidate_tokens_prefer_mixed(self):
        """9999999_ABC-1234_SHEET2.pdf → ABC-1234 (mixed beats 7-digit numeric)"""
        result = extract_part_number_from_filename("9999999_ABC-1234_SHEET2.pdf")
        assert result == "ABC-1234"

    def test_long_numeric_part_number(self):
        """1234567890_SHEET1.pdf → 1234567890 (>6 digits survives filter)"""
        result = extract_part_number_from_filename("1234567890_SHEET1.pdf")
        assert result == "1234567890"

    def test_spaces_as_delimiters(self):
        """33956 MA-S9160-1000 SHEET1.pdf → MA-S9160-1000"""
        result = extract_part_number_from_filename("33956 MA-S9160-1000 SHEET1.pdf")
        assert result == "MA-S9160-1000"

    def test_multiple_mixed_tokens_prefer_longest(self):
        """AB-12_CD-3456_REV_B.pdf → CD-3456 (longer mixed token wins)"""
        result = extract_part_number_from_filename("AB-12_CD-3456_REV_B.pdf")
        assert result == "CD-3456"

    def test_rafael_style_filename(self):
        """REFA-12345-678.pdf → REFA-12345-678"""
        result = extract_part_number_from_filename("REFA-12345-678.pdf")
        assert result == "REFA-12345-678"

    def test_iai_drawing_format(self):
        """IAI_7J14201-100_REV_C.pdf → 7J14201-100"""
        result = extract_part_number_from_filename("IAI_7J14201-100_REV_C.pdf")
        assert result == "7J14201-100"

    # --- Negative / edge cases ---

    def test_only_date_and_sheet(self):
        """SHEET1_04112025.pdf → None (no valid part-number token)"""
        result = extract_part_number_from_filename("SHEET1_04112025.pdf")
        assert result is None

    def test_empty_stem(self):
        """.pdf → None"""
        result = extract_part_number_from_filename(".pdf")
        assert result is None

    def test_only_short_numbers(self):
        """1234.pdf → None (≤4 pure digits filtered), 12345.pdf → '12345' (5+ digits allowed)"""
        result = extract_part_number_from_filename("1234.pdf")
        assert result is None
        result5 = extract_part_number_from_filename("12345.pdf")
        assert result5 == "12345"

    def test_only_single_letter(self):
        """A.pdf → None"""
        result = extract_part_number_from_filename("A.pdf")
        assert result is None

    def test_rev_only(self):
        """REVA.pdf → None (matches REV pattern)"""
        result = extract_part_number_from_filename("REVA.pdf")
        assert result is None

    def test_dwg_only(self):
        """DWG.pdf → None (matches filler word)"""
        result = extract_part_number_from_filename("DWG.pdf")
        assert result is None

    def test_returns_string_type(self):
        result = extract_part_number_from_filename("33956_XY-789_SHEET1.pdf")
        assert isinstance(result, str)


# =====================================================================
# Fallback chain integration  –  simulate result_data scenarios
# =====================================================================

class TestFallbackChainLogic:
    """
    Test the fallback priority logic as implemented in
    customer_extractor_v3_dual.py (extracted here as a pure function
    for deterministic testing).
    """

    @staticmethod
    def _apply_free_fallback(result_data: dict, file_path: str) -> dict:
        """
        Replicate the exact fallback logic from customer_extractor_v3_dual.py
        so we can unit-test it without invoking the full pipeline.
        """
        if not result_data.get('part_number'):
            fallback_pn = None
            fallback_source = None

            # Priority 1: PL main part number
            pl_pn = result_data.get('pl_main_part_number', '')
            if pl_pn and pl_pn != 'MULTIPLE':
                fallback_pn = pl_pn
                fallback_source = 'PL Part Number'

            # Priority 2: Filename extraction
            if not fallback_pn:
                fn_pn = extract_part_number_from_filename(file_path)
                if fn_pn:
                    fallback_pn = fn_pn
                    fallback_source = 'filename extraction'

            if fallback_pn:
                result_data['part_number'] = fallback_pn
                result_data['_fallback_source'] = fallback_source
                if not result_data.get('drawing_number'):
                    result_data['drawing_number'] = fallback_pn

        return result_data

    def test_pl_pn_takes_priority_over_filename(self):
        """PL Part Number should be used before filename extraction."""
        data = {'part_number': '', 'drawing_number': '', 'pl_main_part_number': 'PL-999'}
        result = self._apply_free_fallback(data, '33956_MA-S9160-1000_SHEET1.pdf')
        assert result['part_number'] == 'PL-999'
        assert result['drawing_number'] == 'PL-999'
        assert result['_fallback_source'] == 'PL Part Number'

    def test_filename_used_when_no_pl(self):
        """Filename extraction used when PL part number is absent."""
        data = {'part_number': '', 'drawing_number': ''}
        result = self._apply_free_fallback(data, '33956_MA-S9160-1000_SHEET1.pdf')
        assert result['part_number'] == 'MA-S9160-1000'
        assert result['drawing_number'] == 'MA-S9160-1000'
        assert result['_fallback_source'] == 'filename extraction'

    def test_multiple_pl_pn_skipped(self):
        """PL Part Number = 'MULTIPLE' should be skipped, fall to filename."""
        data = {'part_number': None, 'drawing_number': None, 'pl_main_part_number': 'MULTIPLE'}
        result = self._apply_free_fallback(data, '33956_MA-S9160-1000_SHEET1.pdf')
        assert result['part_number'] == 'MA-S9160-1000'
        assert result['_fallback_source'] == 'filename extraction'

    def test_no_fallback_when_part_exists(self):
        """Fallback should NOT overwrite an existing part_number."""
        data = {'part_number': 'EXISTING-123', 'drawing_number': 'DWG-456', 'pl_main_part_number': 'PL-999'}
        result = self._apply_free_fallback(data, '33956_MA-S9160-1000_SHEET1.pdf')
        assert result['part_number'] == 'EXISTING-123'
        assert '_fallback_source' not in result

    def test_drawing_number_preserved_when_exists(self):
        """If drawing_number already set, only fill part_number."""
        data = {'part_number': '', 'drawing_number': 'DWG-456'}
        result = self._apply_free_fallback(data, '33956_MA-S9160-1000_SHEET1.pdf')
        assert result['part_number'] == 'MA-S9160-1000'
        assert result['drawing_number'] == 'DWG-456'  # unchanged

    def test_all_fallbacks_exhausted(self):
        """When no PL and filename has no valid token, nothing changes."""
        data = {'part_number': '', 'drawing_number': ''}
        result = self._apply_free_fallback(data, 'SHEET1_04112025.pdf')
        assert result['part_number'] == ''
        assert result['drawing_number'] == ''
        assert '_fallback_source' not in result

    def test_none_part_number_triggers_fallback(self):
        """None (not just empty string) should also trigger fallback."""
        data = {'part_number': None, 'drawing_number': None}
        result = self._apply_free_fallback(data, 'XY-7890-100.pdf')
        assert result['part_number'] == 'XY-7890-100'
