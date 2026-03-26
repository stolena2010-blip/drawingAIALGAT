"""Unit tests for src/services/extraction/filename_utils.py"""

from src.services.extraction.filename_utils import (
    check_value_in_filename, check_exact_match_in_filename,
    fix_zero_o_from_filename, _extract_item_number_from_filename,
    _normalize_item_number, _fuzzy_char_equal, _fuzzy_substring_match,
    _get_confusion_positions, _generate_candidates,
    _score_candidate_against_filename, _disambiguate_part_number,
)


# =========================================================================
# A) check_value_in_filename
# =========================================================================

def test_check_value_found():
    assert check_value_in_filename("3585410", "MDMD_3585410_30.pdf") == True

def test_check_value_not_found():
    assert check_value_in_filename("9999999", "MDMD_3585410_30.pdf") == False

def test_check_value_none():
    assert check_value_in_filename(None, "some_file.pdf") == False

def test_check_value_null_string():
    assert check_value_in_filename("null", "some_file.pdf") == False

def test_check_value_too_short():
    assert check_value_in_filename("AB", "AB_file.pdf") == False

def test_check_value_case_insensitive():
    assert check_value_in_filename("ABC123", "abc123_drawing.pdf") == True

def test_check_value_ignores_separators():
    assert check_value_in_filename("ABC123", "ABC-123_drawing.pdf") == True


# =========================================================================
# B) check_exact_match_in_filename
# =========================================================================

def test_exact_match_true():
    # Normalized filename includes "pdf", so only bare name matches exactly
    assert check_exact_match_in_filename("3585410", "3585410") == True

def test_exact_match_with_extension_false():
    # ".pdf" merges into normalized string -> "3585410pdf" != "3585410"
    assert check_exact_match_in_filename("3585410", "3585410.pdf") == False

def test_exact_match_with_prefix():
    # Normalized: "mdmd358541030pdf" — value not bounded
    assert check_exact_match_in_filename("3585410", "MDMD_3585410_30.pdf") == False

def test_exact_match_partial_false():
    assert check_exact_match_in_filename("35854", "3585410.pdf") == False

def test_exact_match_none():
    assert check_exact_match_in_filename(None, "file.pdf") == False


# =========================================================================
# C) fix_zero_o_from_filename
# =========================================================================

def test_fix_zero_o_corrects():
    # OCR reads "O" but filename has "0"
    result = fix_zero_o_from_filename("BO52931A", "B052931A_30.pdf")
    assert result == "B052931A"

def test_fix_zero_o_no_change_needed():
    result = fix_zero_o_from_filename("B052931A", "B052931A_30.pdf")
    assert result == "B052931A"

def test_fix_zero_o_none_value():
    assert fix_zero_o_from_filename(None, "file.pdf") is None

def test_fix_zero_o_empty_value():
    assert fix_zero_o_from_filename("", "file.pdf") == ""


# =========================================================================
# D) _extract_item_number_from_filename
# =========================================================================

def test_extract_item_mixed_alphanumeric():
    assert _extract_item_number_from_filename("68A250781_30.pdf") == "68a250781"

def test_extract_item_with_prefix():
    assert _extract_item_number_from_filename("MDMD_3585410_30.pdf") == "3585410"

def test_extract_item_pure_number():
    assert _extract_item_number_from_filename("3585410.pdf") == "3585410"

def test_extract_item_empty():
    assert _extract_item_number_from_filename("") == ""

def test_extract_item_removes_pl_suffix():
    result = _extract_item_number_from_filename("3585410_PL.pdf")
    assert "pl" not in result.lower()


# =========================================================================
# E) _normalize_item_number
# =========================================================================

def test_normalize_removes_brackets():
    assert _normalize_item_number("BO52931A [D]") == "b052931a"

def test_normalize_removes_trailing_zeros():
    result = _normalize_item_number("B052931A-000")
    assert result == "b052931a"

def test_normalize_removes_separators():
    result = _normalize_item_number("B-052-931A")
    assert "b052931a" == result

def test_normalize_empty():
    assert _normalize_item_number("") == ""

def test_normalize_none():
    assert _normalize_item_number(None) == ""


# =========================================================================
# F) _fuzzy_char_equal
# =========================================================================

def test_fuzzy_same_char():
    assert _fuzzy_char_equal("A", "A") == True

def test_fuzzy_O_zero():
    assert _fuzzy_char_equal("O", "0") == True

def test_fuzzy_zero_O():
    assert _fuzzy_char_equal("0", "O") == True

def test_fuzzy_I_one():
    assert _fuzzy_char_equal("I", "1") == True

def test_fuzzy_l_one():
    assert _fuzzy_char_equal("l", "1") == True

def test_fuzzy_different():
    assert _fuzzy_char_equal("A", "B") == False


# =========================================================================
# G) _fuzzy_substring_match
# =========================================================================

def test_fuzzy_match_exact():
    assert _fuzzy_substring_match("35854", "35854_drawing.pdf") == True

def test_fuzzy_match_with_ocr_confusion():
    # "O" in part_num matches "0" in filename
    assert _fuzzy_substring_match("B0529O1A", "B052901A_30.pdf") == True

def test_fuzzy_match_too_short():
    assert _fuzzy_substring_match("AB", "AB_file.pdf", min_length=5) == False

def test_fuzzy_match_not_found():
    assert _fuzzy_substring_match("ZZZZZ", "3585410.pdf") == False
