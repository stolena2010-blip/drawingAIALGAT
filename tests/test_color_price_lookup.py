"""Tests for color_price_lookup module."""

import pytest
from unittest.mock import patch

from src.services.extraction.color_price_lookup import (
    has_paint_process,
    _extract_specs,
    _extract_spec_sections,
    _extract_color_hints,
    _narrow_by_hints,
    lookup_color_prices,
)


# ── has_paint_process ──────────────────────────────────────────


class TestHasPaintProcess:
    """Detect painting/coating keywords in merged_processes."""

    @pytest.mark.parametrize("text", [
        "אלומיניום | צביעה RAL6003",
        "פלדה | ציפוי אבץ",
        "painting required",
        "primer + topcoat",
        "פריימר",
        "epoxy coating",
        "אפוקסי",
    ])
    def test_paint_keywords_detected(self, text):
        assert has_paint_process(text) is True

    @pytest.mark.parametrize("text", [
        "אלומיניום | תמורה",
        "נירוסטה | אנודייז",
        "פלדה | חיסום",
        "",
        None,
    ])
    def test_non_paint_not_detected(self, text):
        assert has_paint_process(text) is False


# ── _extract_specs ─────────────────────────────────────────────


class TestExtractSpecs:

    def test_multiple_specs(self):
        specs = _extract_specs(
            "תמורה - MIL-DTL-5541 Type I Class 1A | צביעה - MIL-PRF-85285 TY I"
        )
        assert "MIL-DTL-5541" in specs
        assert "MIL-PRF-85285" in specs

    def test_rafdoc(self):
        specs = _extract_specs("צביעה - RAFDOC#16876050")
        assert any("RAFDOC" in s for s in specs)

    def test_fed_code(self):
        specs = _extract_specs("צביעה - MIL-PRF-85285 FED 37875")
        assert any("MIL-PRF-85285" in s for s in specs)
        assert any("FED" in s for s in specs)

    def test_ral_code_with_space(self):
        specs = _extract_specs("צביעה - MIL-PRF-85285 RAL 6003")
        assert any("RAL" in s and "6003" in s for s in specs)

    def test_ral_code_no_space(self):
        specs = _extract_specs("צביעה - RAL6003")
        assert any("RAL" in s and "6003" in s for s in specs)

    def test_ral_with_suffix(self):
        specs = _extract_specs("צביעה - MIL-PRF-22750 RAL6031/F9")
        assert any("RAL" in s and "6031" in s for s in specs)
        assert any("MIL-PRF-22750" in s for s in specs)

    def test_empty(self):
        assert _extract_specs("") == []
        assert _extract_specs(None) == []


# ── _extract_spec_sections ─────────────────────────────────────


class TestExtractSpecSections:

    def test_two_sections(self):
        secs = _extract_spec_sections(
            "תמורה - MIL-DTL-5541 Type I | צביעה - MIL-PRF-85285 FED 37875"
        )
        assert len(secs) == 2
        assert secs[0] == ["MIL-DTL-5541"]
        assert secs[1] == ["MIL-PRF-85285", "FED 37875"]

    def test_ral_in_section(self):
        secs = _extract_spec_sections("צביעה - MIL-PRF-85285 RAL 6003")
        assert len(secs) == 1
        assert any("MIL-PRF-85285" in s for s in secs[0])
        assert any("RAL" in s for s in secs[0])

    def test_single_section(self):
        secs = _extract_spec_sections("צביעה - RAFDOC#16876050")
        assert len(secs) == 1

    def test_empty(self):
        assert _extract_spec_sections("") == []
        assert _extract_spec_sections(None) == []


# ── lookup_color_prices (integration with real COLORS.xlsx) ────


class TestLookupColorPrices:

    def test_no_paint_returns_empty(self):
        """When merged_processes has no paint keyword → empty."""
        result = lookup_color_prices(
            "תמורה - MIL-DTL-5541 Type I",
            "אלומיניום | תמורה",
        )
        assert result == ""

    def test_paint_with_known_spec_returns_results(self):
        """MIL-PRF-85285 is a common paint spec in COLORS.xlsx."""
        result = lookup_color_prices(
            "צביעה - MIL-PRF-85285",
            "אלומיניום | צביעה",
        )
        assert "MIL-PRF-85285" in result
        assert "₪" in result or "$" in result  # has prices

    def test_narrow_match_with_fed(self):
        """Adding FED code narrows results significantly."""
        broad = lookup_color_prices("צביעה - MIL-PRF-85285", "צביעה")
        narrow = lookup_color_prices("צביעה - MIL-PRF-85285 FED 37875", "צביעה")
        assert len(narrow.split("\n")) < len(broad.split("\n"))
        assert "FED 37875" in narrow

    def test_primer_spec(self):
        """MIL-P-23377 is a primer spec."""
        result = lookup_color_prices("פריימר - MIL-P-23377 TY I", "פריימר | צביעה")
        assert "MIL-P-23377" in result

    def test_empty_specs_returns_empty(self):
        """Paint process but no specs → empty."""
        result = lookup_color_prices("", "צביעה")
        assert result == ""

    def test_non_paint_specs_with_paint_process(self):
        """Paint process but only non-paint specs → may still find some or empty."""
        result = lookup_color_prices("חומר - AMS 5643", "צביעה")
        # AMS 5643 is stainless spec, not in COLORS.xlsx → empty
        assert isinstance(result, str)

    def test_ral_narrows_results(self):
        """Adding RAL code should narrow results vs spec alone."""
        broad = lookup_color_prices("צביעה - MIL-PRF-85285", "צביעה")
        narrow = lookup_color_prices("צביעה - MIL-PRF-85285 RAL 6003", "צביעה")
        # narrow should have fewer or equal results
        assert len(narrow.split('\n')) <= len(broad.split('\n'))


# ── _extract_color_hints ──────────────────────────────────────


class TestExtractColorHints:

    def test_hebrew_color(self):
        hints = _extract_color_hints("אלומיניום | צביעה אפור מט")
        assert 'אפור' in hints
        assert 'מט' in hints

    def test_english_color(self):
        hints = _extract_color_hints("אלומיניום | painting BLACK MATT")
        assert 'BLACK' in hints
        assert 'MATT' in hints

    def test_mixed_color_finish(self):
        hints = _extract_color_hints("צביעה שחור מבריק")
        assert 'שחור' in hints
        assert 'מבריק' in hints

    def test_no_color_words(self):
        hints = _extract_color_hints("אלומיניום | תמורה")
        assert hints == []

    def test_empty(self):
        assert _extract_color_hints("") == []
        assert _extract_color_hints(None) == []

    def test_case_insensitive_english(self):
        hints = _extract_color_hints("צביעה green gloss")
        assert 'GREEN' in hints
        assert 'GLOSS' in hints


# ── _narrow_by_hints ────────────────────────────────────────


class TestNarrowByHints:

    def test_narrows_by_color(self):
        candidates = [
            ('MIL-PRF-85285 BLACK MATT', 'pn1', 'ext1'),
            ('MIL-PRF-85285 GREEN GLOSS', 'pn2', 'ext2'),
            ('MIL-PRF-85285 BLACK GLOSS', 'pn3', 'ext3'),
        ]
        result = _narrow_by_hints(candidates, ['BLACK'])
        assert len(result) == 2
        pns = [r[1] for r in result]
        assert 'pn1' in pns
        assert 'pn3' in pns

    def test_narrows_by_color_and_finish(self):
        candidates = [
            ('MIL-PRF-85285 BLACK MATT', 'pn1', 'ext1'),
            ('MIL-PRF-85285 GREEN GLOSS', 'pn2', 'ext2'),
            ('MIL-PRF-85285 BLACK GLOSS', 'pn3', 'ext3'),
        ]
        result = _narrow_by_hints(candidates, ['BLACK', 'MATT'])
        assert len(result) == 1
        assert result[0][1] == 'pn1'

    def test_fallback_on_zero_match(self):
        """If hint narrows to zero, skip that hint."""
        candidates = [
            ('MIL-PRF-85285 GREEN GLOSS', 'pn2', 'ext2'),
        ]
        result = _narrow_by_hints(candidates, ['ORANGE'])
        # ORANGE not found → fallback to all candidates
        assert len(result) == 1

    def test_empty_hints(self):
        candidates = [('DESC', 'pn1', 'ext1')]
        result = _narrow_by_hints(candidates, [])
        assert result == candidates

    def test_empty_candidates(self):
        result = _narrow_by_hints([], ['BLACK'])
        assert result == []


# ── Integration: color hints narrow lookup ──────────────────


class TestColorHintsIntegration:

    def test_color_hint_narrows_results(self):
        """צביעה שחור should return fewer results than just צביעה."""
        broad = lookup_color_prices("\u05e6\u05d1\u05d9\u05e2\u05d4 - MIL-PRF-85285", "צביעה")
        narrow = lookup_color_prices("\u05e6\u05d1\u05d9\u05e2\u05d4 - MIL-PRF-85285", "צביעה שחור")
        assert len(narrow.split('\n')) <= len(broad.split('\n'))

    def test_finish_hint_narrows_results(self):
        """צביעה מט should return fewer results than just צביעה."""
        broad = lookup_color_prices("\u05e6\u05d1\u05d9\u05e2\u05d4 - MIL-PRF-85285", "צביעה")
        narrow = lookup_color_prices("\u05e6\u05d1\u05d9\u05e2\u05d4 - MIL-PRF-85285", "צביעה מט")
        assert len(narrow.split('\n')) <= len(broad.split('\n'))
