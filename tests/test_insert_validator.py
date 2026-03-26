"""Tests for src.services.extraction.insert_validator."""

import pytest
from src.services.extraction.insert_validator import _is_real_insert, validate_inserts_hardware


# ── _is_real_insert unit tests ────────────────────────────────────

class TestIsRealInsert:
    """Test individual insert recognition."""

    def test_ms51830_keensert(self):
        assert _is_real_insert({"cat_no": "402050023", "description": "INSERT LIGHTWEIGHT MS51830"})

    def test_helicoil_ma3279(self):
        assert _is_real_insert({"cat_no": "MA3279-154", "description": "Helicoil M4*1.5d -Free"})

    def test_ms21209_helical(self):
        assert _is_real_insert({"cat_no": "402250111", "description": "MS21209C0215 INSERT, HELICAL, 2-56 X 1.5D"})

    def test_ms122_insert(self):
        assert _is_real_insert({"cat_no": "402000017", "description": "MS122078, INSERT, HELICAL, 6-32 X 1D"})

    def test_keensert_keyword(self):
        assert _is_real_insert({"cat_no": "402050089", "description": "קינסרט MS51830-203"})

    def test_ioi_prefix(self):
        assert _is_real_insert({"cat_no": "IOI1K500-001", "description": "INSERT IOI"})

    def test_kn_prefix(self):
        assert _is_real_insert({"cat_no": "KNL1032J", "description": "KNL1032J /MS51830-201L"})

    def test_nas1149_washer(self):
        assert _is_real_insert({"cat_no": "NAS1149", "description": "WASHER"})

    def test_loctite(self):
        assert _is_real_insert({"cat_no": "", "description": "LOCTITE 601 50CC"})

    def test_rafael_402_pn(self):
        assert _is_real_insert({"cat_no": "402000006", "description": "MS 122076"})

    def test_thread_description_m4(self):
        assert _is_real_insert({"cat_no": "", "description": "M4*0.7 1D/MA3279-104"})

    # ── Negative tests: NOT inserts ──

    def test_bp_part_not_insert(self):
        assert not _is_real_insert({"cat_no": "BP41596A", "description": "PLATE UPPER", "qty": "1"})

    def test_bracket_not_insert(self):
        assert not _is_real_insert({"cat_no": "BP36739A", "description": "BRACKET ASSY", "qty": "1"})

    def test_generic_part_not_insert(self):
        assert not _is_real_insert({"cat_no": "BP36695A", "description": "COVER PANEL", "qty": "4"})

    def test_empty_item(self):
        assert not _is_real_insert({"cat_no": "", "description": ""})

    def test_housing_not_insert(self):
        assert not _is_real_insert({"cat_no": "12345678", "description": "HOUSING MAIN"})

    def test_unknown_pn_no_keywords(self):
        assert not _is_real_insert({"cat_no": "XY12345", "description": "COMPONENT A"})


# ── validate_inserts_hardware integration tests ───────────────────

class TestValidateInsertsHardware:
    """Test the full validation filter."""

    def test_empty_list(self):
        assert validate_inserts_hardware([]) == []

    def test_all_real_inserts(self):
        items = [
            {"cat_no": "402050023", "qty": "4", "description": "INSERT MS51830"},
            {"cat_no": "NAS1149", "qty": "8", "description": "WASHER"},
        ]
        result = validate_inserts_hardware(items)
        assert len(result) == 2

    def test_all_fake_inserts(self):
        """Sub-assembly parts should all be dropped."""
        items = [
            {"cat_no": "BP41596A", "qty": "1", "description": "PLATE UPPER"},
            {"cat_no": "BP36739A", "qty": "1", "description": "BRACKET ASSY"},
            {"cat_no": "BP36695A", "qty": "4", "description": "COVER"},
        ]
        result = validate_inserts_hardware(items, part_number="BO54758A")
        assert len(result) == 0

    def test_mixed_real_and_fake(self):
        """Only real inserts should survive."""
        items = [
            {"cat_no": "402050023", "qty": "4", "description": "INSERT MS51830"},
            {"cat_no": "BP41596A", "qty": "1", "description": "PLATE UPPER"},
            {"cat_no": "MA3279-154", "qty": "2", "description": "Helicoil M4*1.5d"},
        ]
        result = validate_inserts_hardware(items, part_number="TEST")
        assert len(result) == 2
        assert result[0]["cat_no"] == "402050023"
        assert result[1]["cat_no"] == "MA3279-154"

    def test_none_input(self):
        assert validate_inserts_hardware(None) is None

    def test_non_dict_items_skipped(self):
        items = [
            {"cat_no": "402050023", "qty": "4", "description": "INSERT MS51830"},
            "garbage string",
            42,
        ]
        result = validate_inserts_hardware(items)
        assert len(result) == 1
