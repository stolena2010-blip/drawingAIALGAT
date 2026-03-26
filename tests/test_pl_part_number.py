"""Tests for PL main part number extraction — pure logic, no API."""
import pytest
from src.services.reporting.pl_generator import (
    _detect_pl_format,
    _extract_header_part_number,
    _has_suffix,
    _extract_manufactured_items_from_text,
    _determine_pl_main_part_number,
)


class TestDetectPLFormat:
    def test_mantra_golan(self):
        assert _detect_pl_format("PL ID: GST421152\nsome text") == "MANTRA_GOLAN"

    def test_mantra_golan_with_space(self):
        assert _detect_pl_format("PL ID : GST421152\nsome text") == "MANTRA_GOLAN"

    def test_mantra_kbm(self):
        assert _detect_pl_format("Doc ID: UCP-213-16884\nsome text") == "MANTRA_KBM"

    def test_mantra_systems(self):
        assert _detect_pl_format("Doc ID: MD-4341-1016\nsome text") == "MANTRA_SYSTEMS"

    def test_new_tabular(self):
        text = "Part Number Rev Catalog No Description Release Status\n1096G454-001"
        assert _detect_pl_format(text) == "NEW_TABULAR"

    def test_unknown(self):
        assert _detect_pl_format("random text with no PL markers") == "UNKNOWN"


class TestExtractHeaderPartNumber:
    def test_golan_pl_id(self):
        text = "PL ID: GST421152\nother stuff"
        assert _extract_header_part_number(text, "MANTRA_GOLAN") == "GST421152"

    def test_systems_doc_id(self):
        text = "Doc ID: MD-4341-1016\nother stuff"
        assert _extract_header_part_number(text, "MANTRA_SYSTEMS") == "MD-4341-1016"

    def test_kbm_doc_id(self):
        text = "Doc ID: UCP-213-16884\nother stuff"
        assert _extract_header_part_number(text, "MANTRA_KBM") == "UCP-213-16884"

    def test_tabular_header(self):
        text = "Part Number Rev Catalog No\n1096G454-001 A 12345"
        result = _extract_header_part_number(text, "NEW_TABULAR")
        assert result == "1096G454-001"


class TestHasSuffix:
    def test_with_suffix(self):
        assert _has_suffix("GST421152-001") is True

    def test_without_suffix(self):
        assert _has_suffix("GST421152") is False

    def test_short_suffix(self):
        assert _has_suffix("PART-01") is False  # Only 2 digits, not 3


class TestExtractManufacturedItems:
    def test_golan_cm_items(self):
        text = "PL ID: GST421152\n001 GST421152-001 CM BRACKET\n002 BOLT-123 HM BOLT"
        items = _extract_manufactured_items_from_text(text, "MANTRA_GOLAN")
        assert items == ["GST421152-001"]

    def test_golan_multiple_cm(self):
        text = "001 GST421152-001 CM BRACKET\n003 GST421152-003 CM PLATE\n005 BOLT HM BOLT"
        items = _extract_manufactured_items_from_text(text, "MANTRA_GOLAN")
        assert len(items) == 2
        assert "GST421152-001" in items
        assert "GST421152-003" in items

    def test_skip_raw_materials_901(self):
        text = "001 PART-001 CM BRACKET\n901 RAW-MAT HM ALUMINUM"
        items = _extract_manufactured_items_from_text(text, "MANTRA_GOLAN")
        assert len(items) == 1
        assert items[0] == "PART-001"

    def test_kbm_im_items(self):
        text = "001 UCP-213-16884-601 IM BRACKET\n002 BOLT H BOLT"
        items = _extract_manufactured_items_from_text(text, "MANTRA_KBM")
        assert items == ["UCP-213-16884-601"]

    def test_kbm_deleted_items_skipped(self):
        text = "001 UCP-OLD-001 IM *DELETED PART*\n002 UCP-NEW-001 IN BRACKET"
        items = _extract_manufactured_items_from_text(text, "MANTRA_KBM")
        assert len(items) == 1
        assert items[0] == "UCP-NEW-001"

    def test_systems_mp_items(self):
        text = "001 4341-1016-001 MP BRACKET\n002 BOLT-A HM BOLT"
        items = _extract_manufactured_items_from_text(text, "MANTRA_SYSTEMS")
        assert items == ["4341-1016-001"]


class TestDeterminePLMainPartNumber:
    def test_golan_single_cm(self):
        text = "PL ID: GST421152\n001 GST421152-001 CM BRACKET"
        assert _determine_pl_main_part_number(text) == "GST421152-001"

    def test_golan_multiple_cm(self):
        text = "PL ID: GST421152\n001 GST421152-001 CM BRACKET\n003 GST421152-003 CM PLATE"
        assert _determine_pl_main_part_number(text) == "MULTIPLE"

    def test_golan_no_manufactured_adds_suffix(self):
        text = "PL ID: EFR1120708\n001 BOLT-123 HM BOLT"
        assert _determine_pl_main_part_number(text) == "EFR1120708-001"

    def test_systems_single_mp(self):
        text = "Doc ID: MD-4341-1016\n001 4341-1016-001 MP BRACKET"
        assert _determine_pl_main_part_number(text) == "4341-1016-001"

    def test_systems_no_manufactured(self):
        text = "Doc ID: MD-4350-1118\n901 RAW HM ALUMINUM"
        assert _determine_pl_main_part_number(text) == "MD-4350-1118-001"

    def test_kbm_single_im(self):
        text = "Doc ID: UCP-213-16884\n001 UCP-213-16884-601 IM BRACKET"
        assert _determine_pl_main_part_number(text) == "UCP-213-16884-601"

    def test_tabular_no_make_items(self):
        text = "Part Number Rev Catalog No Description Release Status\n1096G454-001 A 12345 BRACKET Released\n1 BOLT-123 123 Hardware 10 EA Buy"
        assert _determine_pl_main_part_number(text) == "1096G454-001"

    def test_empty_text(self):
        assert _determine_pl_main_part_number("") == ""

    def test_unknown_format(self):
        assert _determine_pl_main_part_number("random text") == ""

    def test_header_already_has_suffix(self):
        # NEW_TABULAR: header already has -001, don't double-append
        text = "Part Number Rev Catalog No Description Release Status\n5903U496-001 B 12345 DOC Released"
        result = _determine_pl_main_part_number(text)
        assert result == "5903U496-001"
        assert not result.endswith("-001-001")  # No double suffix
