"""Tests for src/services/extraction/sanity_checks.py

Covers the three public / testable surfaces:
  - is_cage_code()
  - _find_near_match_in_filename()
  - run_pn_sanity_checks() — selected high-value scenarios
"""
import pytest
from src.services.extraction.sanity_checks import (
    is_cage_code,
    _find_near_match_in_filename,
    run_pn_sanity_checks,
)


# ── is_cage_code ──────────────────────────────────────────────────────

class TestIsCageCode:

    def test_valid_5_char(self):
        assert is_cage_code("K3J40") is True

    def test_valid_6_char(self):
        assert is_cage_code("A1B2C3") is True

    def test_pure_digits_not_cage(self):
        # Must have both letter AND digit — all-digit is not a CAGE code
        assert is_cage_code("12345") is False

    def test_pure_letters_not_cage(self):
        assert is_cage_code("ABCDE") is False

    def test_too_short(self):
        assert is_cage_code("A1B") is False

    def test_too_long(self):
        assert is_cage_code("A1B2C3D") is False

    def test_none_input(self):
        assert is_cage_code(None) is False

    def test_empty_string(self):
        assert is_cage_code("") is False

    def test_typical_part_number_not_cage(self):
        # Real part numbers are longer and therefore not CAGE codes
        assert is_cage_code("FT-15912029-00") is False


# ── _find_near_match_in_filename ──────────────────────────────────────

class TestFindNearMatchInFilename:

    def test_one_char_digit_swap(self):
        # OCR reads '6' instead of '5'.
        # Filename without extension suffix so the compound token is exactly the same length.
        match = _find_near_match_in_filename("FT-16912029", "FT-15912029")
        assert match is not None

    def test_exact_match_returns_none(self):
        # Exact match means no near-match correction needed
        result = _find_near_match_in_filename("FT-15912029", "FT-15912029-00_RevA")
        # An exact match won't be caught by the 1-char-diff logic (diff_count == 0)
        assert result is None

    def test_two_char_diff_not_returned(self):
        # Two-character difference is too far — should not be returned
        result = _find_near_match_in_filename("FT-15902029", "FT-15912129_Rev1")
        assert result is None

    def test_short_value_skipped(self):
        # Values under 4 normalised chars are skipped
        result = _find_near_match_in_filename("AB1", "AB2_drawing")
        assert result is None

    def test_empty_filename(self):
        result = _find_near_match_in_filename("FT-15912029", "")
        assert result is None

    def test_none_value(self):
        result = _find_near_match_in_filename(None, "FT-15912029_RevA")
        assert result is None


# ── run_pn_sanity_checks ──────────────────────────────────────────────

class TestRunPnSanityChecks:

    def _base(self, **overrides):
        data = {
            "part_number": None,
            "drawing_number": None,
            "revision": "",
        }
        data.update(overrides)
        return data

    # ── date rejection ────────────────────────────────────────────────

    def test_date_as_pn_cleared(self):
        # Date is cleared early; the fallback then copies drawing_number to fill part_number.
        # The key requirement: the original date value must NOT survive.
        data = self._base(part_number="01.02.2024", drawing_number="ABC-12345")
        result = run_pn_sanity_checks(data, "ABC-12345_RevA.pdf", "ABC-12345_RevA.pdf")
        assert result["part_number"] != "01.02.2024"

    def test_date_as_dn_cleared(self):
        # Same: date in drawing_number is removed; part_number is unaffected.
        data = self._base(part_number="ABC-12345", drawing_number="2024-03-15")
        result = run_pn_sanity_checks(data, "ABC-12345_RevA.pdf", "ABC-12345_RevA.pdf")
        assert result["drawing_number"] != "2024-03-15"

    def test_valid_pn_not_cleared(self):
        data = self._base(part_number="FT-15912029", drawing_number="FT-15912029")
        result = run_pn_sanity_checks(data, "FT-15912029_RevA.pdf", "FT-15912029_RevA.pdf")
        assert result["part_number"] == "FT-15912029"

    # ── CAGE code removal ─────────────────────────────────────────────

    def test_cage_code_pn_not_in_filename_removed(self):
        # K3J40 is a valid CAGE code and NOT in the filename → removed.
        # The fallback then fills part_number from drawing_number, so the field
        # won’t be None — but the cage-code value must not survive.
        data = self._base(part_number="K3J40", drawing_number="DWG-001")
        result = run_pn_sanity_checks(data, "DWG-001_Rev1.pdf", "DWG-001_Rev1.pdf")
        assert result["part_number"] != "K3J40"

    def test_cage_code_in_filename_kept(self):
        # If the CAGE code IS in the filename, keep it
        data = self._base(part_number="K3J40", drawing_number="K3J40")
        result = run_pn_sanity_checks(data, "K3J40_SomePart.pdf", "K3J40_SomePart.pdf")
        # Should survive since it appears in the filename
        assert result["part_number"] is not None

    # ── REV suffix cleanup ────────────────────────────────────────────

    def test_rev_suffix_stripped_from_pn(self):
        data = self._base(part_number="FT-15912029-00-REVA", drawing_number="FT-15912029-00")
        result = run_pn_sanity_checks(
            data,
            "FT-15912029-00_RevA.pdf",
            "FT-15912029-00_RevA.pdf",
        )
        assert "REV" not in result.get("part_number", "").upper()
        assert result.get("part_number") == "FT-15912029-00"

    def test_rev_suffix_sets_revision_when_empty(self):
        data = self._base(part_number="ABC-123-REVA", drawing_number="ABC-123", revision="")
        result = run_pn_sanity_checks(data, "ABC-123.pdf", "ABC-123.pdf")
        assert result.get("revision", "") == "A"

    # ── fallback copy between fields ─────────────────────────────────

    def test_pn_copied_from_dn_when_pn_missing(self):
        data = self._base(part_number=None, drawing_number="DWG-98765")
        result = run_pn_sanity_checks(data, "DWG-98765_Rev1.pdf", "DWG-98765_Rev1.pdf")
        assert result["part_number"] == "DWG-98765"

    def test_dn_copied_from_pn_when_dn_missing(self):
        data = self._base(part_number="PN-54321", drawing_number=None)
        result = run_pn_sanity_checks(data, "PN-54321.pdf", "PN-54321.pdf")
        assert result["drawing_number"] == "PN-54321"

    # ── IAI unification rule ──────────────────────────────────────────

    def test_iai_sets_dn_from_pn(self):
        data = self._base(part_number="IAI-001", drawing_number=None)
        result = run_pn_sanity_checks(data, "IAI-001.pdf", "IAI-001.pdf", is_iai=True)
        assert result["drawing_number"] == "IAI-001"

    def test_iai_sets_pn_from_dn(self):
        data = self._base(part_number=None, drawing_number="IAI-002")
        result = run_pn_sanity_checks(data, "IAI-002.pdf", "IAI-002.pdf", is_iai=True)
        assert result["part_number"] == "IAI-002"

    # ── sanity check E: DN not in filename, PN is ────────────────────

    def test_dn_replaced_by_pn_when_only_pn_in_filename(self):
        # Non-RAFAEL: DN not in filename, PN is → DN should become PN
        data = self._base(part_number="CORRECT-123", drawing_number="WRONG-999")
        result = run_pn_sanity_checks(
            data,
            "CORRECT-123_Rev1.pdf",
            "CORRECT-123_Rev1.pdf",
            is_rafael=False,
        )
        assert result["drawing_number"] == "CORRECT-123"
