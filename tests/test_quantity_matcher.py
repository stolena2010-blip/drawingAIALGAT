"""Tests for quantity_matcher — per-key filtering & helper logic."""
import pytest
from unittest.mock import patch

from src.services.extraction.quantity_matcher import (
    _key_matches_any_drawing,
    match_quantities_to_drawings,
)
from src.services.extraction.filename_utils import _normalize_item_number


# ─── _key_matches_any_drawing ────────────────────────────────────────

class TestKeyMatchesAnyDrawing:
    """Unit tests for the shared matching helper."""

    def test_exact_match(self):
        assert _key_matches_any_drawing("pa13804324", {"pa13804324"})

    def test_containment_key_in_drawing(self):
        # drawing contains the key
        assert _key_matches_any_drawing("abc", {"xyzabcdef"})

    def test_containment_drawing_in_key(self):
        # email key contains the drawing  (e.g. pal3804324001 ⊃ pal3804324)
        dn = _normalize_item_number("PAL3804324")
        eq = _normalize_item_number("pal3804324001")
        assert _key_matches_any_drawing(eq, {dn})

    def test_suffix_digit_match(self):
        # last 5 digits the same despite different prefix
        assert _key_matches_any_drawing("x12345", {"y12345"})

    def test_suffix_digit_no_match(self):
        assert not _key_matches_any_drawing("x12345", {"y99999"})

    def test_phone_number_no_match(self):
        # "3916" has only 4 digits — shouldn't match a real PN like pa13804324
        assert not _key_matches_any_drawing("3916", {"pa13804324"})

    def test_empty_key_returns_false(self):
        assert not _key_matches_any_drawing("", {"pa13804324"})

    def test_case_insensitive_via_normalize(self):
        """PAL and pal should both normalise identically and match."""
        n1 = _normalize_item_number("PAL3804324")
        n2 = _normalize_item_number("pal3804324")
        assert n1 == n2
        assert _key_matches_any_drawing(n1, {n2})


# ─── Per-key filtering integration ──────────────────────────────────

class TestPerKeyFiltering:
    """Verify that match_quantities_to_drawings drops individual garbage
    keys from email part_quantities while keeping legitimate ones."""

    @staticmethod
    def _make_result(pn: str) -> dict:
        return {"part_number": pn, "item_name": "test"}

    def test_phone_number_dropped_real_kept(self):
        """Phone number '3916' should be dropped; real PN kept."""
        results = [self._make_result("BM05889A")]
        email_data = {
            "part_quantities": {
                "BM05889A": "2",
                "3916": "1",         # phone fragment — garbage
            }
        }
        match_quantities_to_drawings(
            results, {}, email_data, []
        )
        # After the call the garbage key should be gone
        assert "3916" not in email_data["part_quantities"]
        assert "BM05889A" in email_data["part_quantities"]

    def test_all_garbage_clears_entirely(self):
        """If every key is garbage, part_quantities becomes empty."""
        results = [self._make_result("ABC12345")]
        email_data = {
            "part_quantities": {
                "3916": "1",
                "m1lprf": "2",
            }
        }
        match_quantities_to_drawings(
            results, {}, email_data, ["5"]
        )
        assert email_data["part_quantities"] == {}

    def test_all_real_keys_kept(self):
        """When every key is legitimate, nothing is dropped."""
        results = [
            self._make_result("EFR1201105"),
            self._make_result("PAL3804324"),
        ]
        email_data = {
            "part_quantities": {
                "EFR1201105-001": "4",
                "PAL3804324-001": "40",
            }
        }
        match_quantities_to_drawings(
            results, {}, email_data, ["2"]
        )
        # Both keys survive
        assert len(email_data["part_quantities"]) == 2

    def test_mixed_garbage_and_real(self):
        """Only garbage keys are removed; real keys untouched."""
        results = [
            self._make_result("PAL3804324"),
        ]
        email_data = {
            "part_quantities": {
                "PAL3804324-001": "40",
                "073 229 3916": "1",   # phone
            }
        }
        match_quantities_to_drawings(
            results, {}, email_data, []
        )
        pq = email_data["part_quantities"]
        assert len(pq) == 1
        assert "PAL3804324-001" in pq
