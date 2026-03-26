"""Tests for _normalize_item_number — pure logic, no API."""
import pytest
from src.services.extraction.filename_utils import _normalize_item_number


class TestNormalizeItemNumber:
    def test_basic(self):
        # Note: O→0 and lowercase means "12345" stays "12345"
        assert _normalize_item_number("12345") == "12345"

    def test_strips_whitespace(self):
        result = _normalize_item_number("  12345  ")
        assert result == "12345"

    def test_removes_dashes(self):
        result = _normalize_item_number("123-45-A")
        assert result  # Should return something non-empty
        assert "-" not in result

    def test_empty_string(self):
        result = _normalize_item_number("")
        assert result == ""

    def test_none_safe(self):
        result = _normalize_item_number(None)
        assert result == ""

    def test_removes_brackets(self):
        # "BO52931A [D]" → normalized without brackets
        result = _normalize_item_number("BO52931A [D]")
        assert "[" not in result
        assert "]" not in result

    def test_removes_parentheses(self):
        result = _normalize_item_number("PART(REV)")
        assert "(" not in result

    def test_ocr_o_to_zero(self):
        # O/o → 0
        r1 = _normalize_item_number("BO52931A")
        r2 = _normalize_item_number("B052931A")
        assert r1 == r2

    def test_ocr_i_to_one(self):
        # I/i/l → 1
        r1 = _normalize_item_number("A1B")
        r2 = _normalize_item_number("AIB")
        assert r1 == r2

    def test_trailing_zeros_stripped(self):
        # Trailing zeros removed for strings > 3 chars
        r1 = _normalize_item_number("B052931A-000")
        r2 = _normalize_item_number("B052931A")
        assert r1 == r2

    def test_lowercase(self):
        result = _normalize_item_number("ABC123")
        # After O→0, I→1, lowercase: "abc123" → "a0c123" (no, wait)
        # Actually: A stays A→lowercase a, B→lowercase b, C→lowercase c
        # But O→0 and I/i/l→1
        assert result == result.lower()

    def test_removes_underscores(self):
        result = _normalize_item_number("Tube_Spacer")
        assert "_" not in result

    def test_case_consistency_L(self):
        """PAL3804324 (uppercase) and pal3804324 (lowercase) must normalise the same."""
        assert _normalize_item_number("PAL3804324") == _normalize_item_number("pal3804324")

    def test_case_consistency_containment(self):
        """pal3804324 must be contained in pal3804324001 after normalisation."""
        base = _normalize_item_number("PAL3804324")
        with_suffix = _normalize_item_number("pal3804324001")
        assert base in with_suffix, f"{base!r} not found in {with_suffix!r}"
