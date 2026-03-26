"""Tests for _build_structured_bom, _format_insert_entry, _count_pl_primary_types, _build_merged_description,
_sum_pl_primary_qty, _sum_drawing_primary_qty, _calc_hardware_count."""

import pytest

from src.services.extraction.stage9_merge import (
    _build_merged_description,
    _build_structured_bom,
    _calc_hardware_count,
    _count_pl_primary_types,
    _format_insert_entry,
    _sum_drawing_primary_qty,
    _sum_pl_primary_qty,
)


# ── _format_insert_entry ───────────────────────────────────────


class TestFormatInsertEntry:
    def test_basic_with_price(self):
        assert _format_insert_entry("K500-001", "12", 0.45, "₪") == "K500-001 ×12 ×0.45₪"

    def test_no_price(self):
        assert _format_insert_entry("MS51835", "4", None, "") == "MS51835 ×4"

    def test_dollar_currency(self):
        assert _format_insert_entry("NAS1149", "8", 0.12, "$") == "NAS1149 ×8 ×0.12$"

    def test_no_qty(self):
        assert _format_insert_entry("ABC123", "", 1.5, "₪") == "ABC123 ×1.5₪"

    def test_empty_cat(self):
        assert _format_insert_entry("", "4", None, "") == ""


# ── _count_pl_primary_types ────────────────────────────────────


class TestCountPLPrimaryTypes:
    def test_two_groups(self):
        assert _count_pl_primary_types("K500-001 ×12 ×0.45₪, MS124655 (חלופי) ×12 | MS51831-203 ×4") == 2

    def test_single_group(self):
        assert _count_pl_primary_types("K500-001 ×4 ×0.45₪") == 1

    def test_three_groups(self):
        assert _count_pl_primary_types("A ×1 | B ×2 | C ×3") == 3

    def test_empty(self):
        assert _count_pl_primary_types("") == 0

    def test_none_coerced(self):
        assert _count_pl_primary_types("") == 0

    def test_alternates_not_counted(self):
        # One group with primary + 2 alternates still counts as 1
        assert _count_pl_primary_types("K500 ×4 ×1₪, ALT1 (חלופי) ×4, ALT2 (חלופי) ×4") == 1


# ── _build_structured_bom ──────────────────────────────────────


class TestBuildStructuredBom:
    def test_both_sources(self):
        item = {
            "inserts_hardware": [
                {"cat_no": "K500-001", "qty": "12", "unit_price": 0.45, "currency": "₪"},
                {"cat_no": "MS51835", "qty": "4", "unit_price": None, "currency": ""},
            ],
            "PL Hardware": "K500-001 ×12 ×0.45₪, MS124655 (חלופי) ×12 | 1191-3CN0190S ×12",
        }
        result = _build_structured_bom(item)
        lines = result.split("\n")
        assert len(lines) == 3
        assert lines[0] == "קשיחים [2]:"
        assert lines[1].startswith("שרטוט: ")
        assert "K500-001 ×12 ×0.45₪" in lines[1]
        assert "MS51835 ×4" in lines[1]
        assert lines[2].startswith("עץ: ")
        assert "(חלופי)" in lines[2]

    def test_drawing_only(self):
        item = {
            "inserts_hardware": [
                {"cat_no": "ABC-1", "qty": "2", "unit_price": 3.0, "currency": "$"},
            ],
        }
        result = _build_structured_bom(item)
        assert result == "קשיחים [1]: ABC-1 ×2 ×3.0$"
        assert "שרטוט" not in result

    def test_pl_only(self):
        item = {
            "PL Hardware": "K500-001 ×4 ×0.45₪ | MS51831-203 ×8 ×7.0₪",
        }
        result = _build_structured_bom(item)
        assert result == "קשיחים [2]: K500-001 ×4 ×0.45₪ | MS51831-203 ×8 ×7.0₪"

    def test_no_inserts(self):
        item = {"part_number": "123-456"}
        assert _build_structured_bom(item) == ""

    def test_empty_hardware_list(self):
        item = {"inserts_hardware": [], "PL Hardware": ""}
        assert _build_structured_bom(item) == ""

    def test_string_inserts_hardware_fallback(self):
        """If inserts_hardware is already a string (edge case), use it as-is."""
        item = {
            "inserts_hardware": "K500-001 ×12",
            "PL Hardware": "MS51835 ×4",
        }
        result = _build_structured_bom(item)
        assert "קשיחים [1]:" in result
        assert "שרטוט: K500-001 ×12" in result
        assert "עץ: MS51835 ×4" in result

    def test_skips_empty_cat_no(self):
        item = {
            "inserts_hardware": [
                {"cat_no": "", "qty": "4", "unit_price": None, "currency": ""},
                {"cat_no": "GOOD-1", "qty": "2", "unit_price": None, "currency": ""},
            ],
        }
        result = _build_structured_bom(item)
        assert "קשיחים [1]: GOOD-1 ×2" in result

    def test_count_discrepancy_pl_wins(self):
        """Drawing has 2 types, PL has 3 → PL count wins."""
        item = {
            "inserts_hardware": [
                {"cat_no": "A", "qty": "1", "unit_price": None, "currency": ""},
                {"cat_no": "B", "qty": "2", "unit_price": None, "currency": ""},
            ],
            "PL Hardware": "A ×1 | B ×2 | C ×3",
        }
        result = _build_structured_bom(item)
        assert result.startswith("קשיחים [3]:")

    def test_count_agreement(self):
        """Drawing and PL both have 2 types → [2]."""
        item = {
            "inserts_hardware": [
                {"cat_no": "A", "qty": "1", "unit_price": None, "currency": ""},
                {"cat_no": "B", "qty": "2", "unit_price": None, "currency": ""},
            ],
            "PL Hardware": "A ×1 | B ×2",
        }
        result = _build_structured_bom(item)
        assert result.startswith("קשיחים [2]:")

    def test_none_cat_no_and_qty_safe(self):
        """None values in cat_no/qty must not crash (was AttributeError)."""
        item = {
            "inserts_hardware": [
                {"cat_no": None, "qty": None, "unit_price": None, "currency": None},
                {"cat_no": "GOOD", "qty": "3", "unit_price": 1.0, "currency": "₪"},
            ],
        }
        result = _build_structured_bom(item)
        assert "GOOD ×3 ×1.0₪" in result
        assert "None" not in result


# ── _sum_pl_primary_qty ────────────────────────────────────────


class TestSumPLPrimaryQty:
    def test_two_primary_groups(self):
        assert _sum_pl_primary_qty("K500 ×4 ×0.45₪ | MS51835 ×8") == 12

    def test_single_group_with_alternate(self):
        assert _sum_pl_primary_qty("K500 ×4 ×0.45₪, ALT1 (חלופי) ×4") == 4

    def test_multiple_groups_complex(self):
        # 6 primary groups from the screenshot example: 4+1+8+2+2+2 = 19
        pl = "402050056 ×4 | 4047536336 ×1 | 4021136443 ×8 | 4047735160 ×2 | 4047731807 ×2 | 4047792011 ×2"
        assert _sum_pl_primary_qty(pl) == 19

    def test_empty(self):
        assert _sum_pl_primary_qty("") == 0

    def test_no_qty(self):
        assert _sum_pl_primary_qty("ABC123") == 0

    def test_alternates_excluded(self):
        pl = "K500 ×4, ALT1 (חלופי) ×4, ALT2 (חלופי) ×3 | MS51835 ×8"
        assert _sum_pl_primary_qty(pl) == 12  # 4 + 8, alternates excluded


# ── _sum_drawing_primary_qty ───────────────────────────────────


class TestSumDrawingPrimaryQty:
    def test_basic(self):
        hw = [
            {"cat_no": "A", "qty": "4"},
            {"cat_no": "B", "qty": "2"},
        ]
        assert _sum_drawing_primary_qty(hw) == 6

    def test_empty_list(self):
        assert _sum_drawing_primary_qty([]) == 0

    def test_not_list(self):
        assert _sum_drawing_primary_qty(None) == 0
        assert _sum_drawing_primary_qty("string") == 0

    def test_skips_empty_cat(self):
        hw = [
            {"cat_no": "", "qty": "4"},
            {"cat_no": "GOOD", "qty": "3"},
        ]
        assert _sum_drawing_primary_qty(hw) == 3

    def test_non_dict_entries(self):
        hw = [None, "bad", {"cat_no": "OK", "qty": "5"}]
        assert _sum_drawing_primary_qty(hw) == 5


# ── _calc_hardware_count ───────────────────────────────────────


class TestCalcHardwareCount:
    def test_pl_wins_over_drawing(self):
        item = {
            "inserts_hardware": [
                {"cat_no": "A", "qty": "10"},
            ],
            "PL Hardware": "X ×4 | Y ×8",
        }
        assert _calc_hardware_count(item) == 12  # PL wins

    def test_drawing_only(self):
        item = {
            "inserts_hardware": [
                {"cat_no": "A", "qty": "3"},
                {"cat_no": "B", "qty": "7"},
            ],
        }
        assert _calc_hardware_count(item) == 10

    def test_pl_only(self):
        item = {
            "PL Hardware": "X ×5 | Y ×3",
        }
        assert _calc_hardware_count(item) == 8

    def test_no_hardware(self):
        item = {"part_number": "123"}
        assert _calc_hardware_count(item) == 0

    def test_empty_hardware(self):
        item = {"inserts_hardware": [], "PL Hardware": ""}
        assert _calc_hardware_count(item) == 0


# ── _build_merged_description ──────────────────────────────────


class TestBuildMergedDescription:
    """Tests for the combined field builder."""

    def test_processes_with_hardware(self):
        item = {
            'merged_processes': 'אלומיניום 5052 | אנודייז',
            'inserts_hardware': [
                {"cat_no": "K500", "qty": "12"},
                {"cat_no": "MS51835", "qty": "4"},
            ],
        }
        result = _build_merged_description(item)
        assert result == '(תהליכים) אלומיניום 5052 | אנודייז | H.C=16'

    def test_processes_no_hardware(self):
        item = {'merged_processes': 'חיתוך'}
        result = _build_merged_description(item)
        assert result == '(תהליכים) חיתוך'

    def test_notes_excluded(self):
        item = {
            'merged_processes': 'כיפוף',
            'merged_notes': 'כללי: דחוף',
            'merged_bom': 'קשיחים [2]: A ×2 | B ×1',
        }
        result = _build_merged_description(item)
        assert '(הערות)' not in result
        assert '(עץ)' not in result
        assert result == '(תהליכים) כיפוף'

    def test_pl_wins_in_description(self):
        item = {
            'merged_processes': 'ציפוי',
            'inserts_hardware': [{"cat_no": "A", "qty": "10"}],
            'PL Hardware': 'X ×4 | Y ×8',
        }
        result = _build_merged_description(item)
        assert result == '(תהליכים) ציפוי | H.C=12'

    def test_only_hardware_no_processes(self):
        item = {
            'merged_processes': '',
            'inserts_hardware': [{"cat_no": "A", "qty": "5"}],
        }
        result = _build_merged_description(item)
        assert result == 'H.C=5'

    def test_empty_item(self):
        result = _build_merged_description({})
        assert result == ''

    def test_nan_values_skipped(self):
        item = {'merged_processes': 'nan', 'merged_bom': 'nan', 'merged_notes': 'nan'}
        result = _build_merged_description(item)
        assert result == ''

    def test_whitespace_only_skipped(self):
        item = {'merged_processes': '  ', 'merged_bom': '', 'merged_notes': '\t'}
        result = _build_merged_description(item)
        assert result == ''

    def test_real_world_example(self):
        """Reproduce the screenshot case: 6 inserts totalling qty 19."""
        item = {
            'merged_processes': 'אלומיניום T6511-6061 | תמורה | סימון דיו שחור MM3',
            'inserts_hardware': [
                {"cat_no": "402050056", "qty": "4", "unit_price": None, "currency": ""},
                {"cat_no": "4047536336", "qty": "1", "unit_price": None, "currency": ""},
                {"cat_no": "4021136443", "qty": "8", "unit_price": None, "currency": ""},
                {"cat_no": "4047735160", "qty": "2", "unit_price": None, "currency": ""},
                {"cat_no": "4047731807", "qty": "2", "unit_price": None, "currency": ""},
                {"cat_no": "4047792011", "qty": "2", "unit_price": None, "currency": ""},
            ],
            'merged_notes': 'כללי: ציפוי, אינסרטים/הליקווילים, סימון לייזר/סימון/הדפסה',
        }
        result = _build_merged_description(item)
        assert result == '(תהליכים) אלומיניום T6511-6061 | תמורה | סימון דיו שחור MM3 | H.C=19'
        # No notes, no עץ
        assert '(הערות)' not in result
        assert '(עץ)' not in result
