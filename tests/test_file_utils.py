"""Unit tests for src/services/file/file_utils.py (pure helpers only)."""

from src.services.file.file_utils import (
    _get_file_metadata, _detect_text_heavy_pdf,
    _build_drawing_part_map,
)

from pathlib import Path

SAMPLE_DIR = Path(__file__).parent.parent / "sample"


# =========================================================================
# _get_file_metadata tests
# =========================================================================

def test_get_file_metadata_returns_dict():
    """Test that _get_file_metadata returns expected structure."""
    sample_files = list(SAMPLE_DIR.glob("*.pdf")) + list(SAMPLE_DIR.glob("*.PDF"))
    if not sample_files:
        import pytest
        pytest.skip("No sample PDFs found")
    result = _get_file_metadata(sample_files[0])
    assert isinstance(result, dict)
    assert "file_name" in result
    assert "file_size_mb" in result
    assert "is_large_file" in result
    assert "is_too_large" in result


def test_get_file_metadata_has_dates():
    """Test that metadata includes date fields."""
    sample_files = list(SAMPLE_DIR.glob("*.pdf")) + list(SAMPLE_DIR.glob("*.PDF"))
    if not sample_files:
        import pytest
        pytest.skip("No sample PDFs found")
    result = _get_file_metadata(sample_files[0])
    assert "created_date" in result
    assert "modified_date" in result


# =========================================================================
# _build_drawing_part_map tests
# =========================================================================

def test_build_drawing_part_map_empty():
    result = _build_drawing_part_map([])
    assert isinstance(result, dict)
    assert len(result) == 0


def test_build_drawing_part_map_no_drawing_results():
    """When drawing_results is None, should return empty map."""
    classifications = [
        {"file_path": Path("3585410_30.pdf"), "file_type": "DRAWING", "item_number": "3585410"},
    ]
    result = _build_drawing_part_map(classifications, drawing_results=None)
    assert isinstance(result, dict)
    assert len(result) == 0


def test_build_drawing_part_map_with_data():
    """When drawing_results are provided, should map item numbers to part numbers."""
    classifications = [
        {"file_path": Path("3585410_30.pdf"), "file_type": "DRAWING"},
        {"file_path": Path("order.pdf"), "file_type": "Order"},
    ]
    drawing_results = [
        {"file_name": "3585410_30.pdf", "part_number": "PN-3585410"},
    ]
    result = _build_drawing_part_map(classifications, drawing_results=drawing_results)
    assert isinstance(result, dict)


# =========================================================================
# _detect_text_heavy_pdf tests
# =========================================================================

def test_detect_text_heavy_pdf_returns_tuple():
    """Test that _detect_text_heavy_pdf returns (bool, str)."""
    sample_files = list(SAMPLE_DIR.glob("*.pdf")) + list(SAMPLE_DIR.glob("*.PDF"))
    if not sample_files:
        import pytest
        pytest.skip("No sample PDFs found")
    result = _detect_text_heavy_pdf(sample_files[0])
    assert isinstance(result, tuple)
    assert len(result) == 2
    assert isinstance(result[0], bool)
    assert isinstance(result[1], str)
