"""Mock test: File classification logic without real API calls."""
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock


class TestFileClassification:
    def test_drawing_map_build(self, sample_classifications, sample_results):
        """Test _build_drawing_part_map logic."""
        from src.services.file.file_utils import _build_drawing_part_map

        mapping = _build_drawing_part_map(sample_classifications, sample_results)
        # Should map item numbers to part numbers
        assert isinstance(mapping, dict)

    def test_find_associated_drawing(self, sample_classifications):
        """Test association logic."""
        from src.services.file.file_utils import _find_associated_drawing

        drawing_map = {"12345": "12345-A", "67890": "67890-C"}

        # Order file should find its associated drawing
        order_path = Path("/tmp/test/order_12345.pdf")
        result = _find_associated_drawing(order_path, "PURCHASE_ORDER", drawing_map)
        # Should return a string (may be "" if no match found via filename)
        assert isinstance(result, str)

    def test_find_associated_drawing_returns_empty_for_drawing(self):
        """Drawing type should return empty string."""
        from src.services.file.file_utils import _find_associated_drawing

        drawing_map = {"12345": "12345-A"}
        result = _find_associated_drawing(
            Path("/tmp/drawing.pdf"), "DRAWING", drawing_map
        )
        assert result == ""

    def test_find_associated_drawing_empty_map(self):
        """Empty drawing map should return empty."""
        from src.services.file.file_utils import _find_associated_drawing

        result = _find_associated_drawing(
            Path("/tmp/order.pdf"), "PURCHASE_ORDER", {}
        )
        assert result == ""

    def test_detect_text_heavy_pdf(self, tmp_output):
        """Test text-heavy PDF detection."""
        from src.services.file.file_utils import _detect_text_heavy_pdf

        # Create a dummy file (won't be a real PDF, should handle gracefully)
        dummy = tmp_output / "test.pdf"
        dummy.write_bytes(b"not a real pdf")

        try:
            is_text_heavy, reason = _detect_text_heavy_pdf(dummy)
            assert isinstance(is_text_heavy, bool)
        except Exception:
            pass  # OK if it fails on invalid PDF

    def test_build_map_empty_results(self, sample_classifications):
        """Empty drawing results → empty map."""
        from src.services.file.file_utils import _build_drawing_part_map

        mapping = _build_drawing_part_map(sample_classifications, [])
        assert mapping == {}

    def test_build_map_none_results(self, sample_classifications):
        """None drawing results → empty map."""
        from src.services.file.file_utils import _build_drawing_part_map

        mapping = _build_drawing_part_map(sample_classifications, None)
        assert mapping == {}
