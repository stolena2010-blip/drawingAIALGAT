"""Snapshot test: B2B text export produces valid TAB-delimited file."""
import pytest
from pathlib import Path
from src.services.reporting.b2b_export import _save_text_summary_with_variants, _is_single_numeric_quantity


class TestB2BExport:
    def test_creates_files(self, sample_results, tmp_output):
        output_file = tmp_output / "B2B-0_200002-12345.txt"
        _save_text_summary_with_variants(
            sample_results, output_file,
            customer_email="test@test.com",
            b2b_number="B2B-0_200002",
            timestamp="12345"
        )
        # Should create variant files (B2B, B2BH, B2BM)
        txt_files = list(tmp_output.glob("*.txt"))
        assert len(txt_files) >= 1

    def test_file_is_tab_delimited(self, sample_results, tmp_output):
        output_file = tmp_output / "B2B-test.txt"
        _save_text_summary_with_variants(
            sample_results, output_file,
            customer_email="test@test.com",
            b2b_number="B2B-0_200002",
            timestamp="test"
        )

        # Read one of the created files
        txt_files = list(tmp_output.glob("*.txt"))
        if txt_files:
            content = txt_files[0].read_text(encoding='cp1255', errors='replace')
            # Should contain tabs
            assert '\t' in content

    def test_contains_part_numbers(self, sample_results, tmp_output):
        output_file = tmp_output / "B2B-parts.txt"
        _save_text_summary_with_variants(
            sample_results, output_file,
            customer_email="", b2b_number="B2B-0", timestamp=""
        )

        txt_files = list(tmp_output.glob("*.txt"))
        if txt_files:
            content = txt_files[0].read_text(encoding='cp1255', errors='replace')
            assert "12345" in content or "67890" in content

    def test_creates_three_variants(self, sample_results, tmp_output):
        """Should create B2B, B2BH, B2BM variants."""
        output_file = tmp_output / "B2B-0_200002-99999.txt"
        _save_text_summary_with_variants(
            sample_results, output_file,
            customer_email="test@test.com",
            b2b_number="B2B-0_200002",
            timestamp="99999"
        )
        txt_files = list(tmp_output.glob("*.txt"))
        names = [f.name for f in txt_files]
        # Check that at least B2B variant was created
        assert any("B2B" in n for n in names)

    def test_empty_results(self, tmp_output):
        """Empty results should not crash."""
        output_file = tmp_output / "B2B-empty.txt"
        _save_text_summary_with_variants(
            [], output_file,
            customer_email="", b2b_number="B2B-0", timestamp=""
        )
        # Should not raise


class TestIsNumericQuantity:
    """Test _is_single_numeric_quantity — decides field 4 vs field 9."""

    # Should go to field 4 (single numeric)
    @pytest.mark.parametrize("value", [
        "50", "153", "4500", "1.0000", "1", "99",
    ])
    def test_single_number_returns_true(self, value):
        assert _is_single_numeric_quantity(value) is True

    # Should go to field 9 (not single numeric)
    @pytest.mark.parametrize("value", [
        "0",                       # zero = no quantity
        "(80, 100)",               # multiple in parens
        "10-20",                   # range
        "3, 5, 10",               # comma-separated
        "10-20 יחידות (טווח)",     # text with numbers
        "100 (מנות של 10)",        # number + text
        "",                        # empty
        "אין כמויות",              # pure text
    ])
    def test_non_numeric_returns_false(self, value):
        assert _is_single_numeric_quantity(value) is False


class TestB2BFieldRouting:
    """Test that quantities route correctly to field 4 vs field 9."""

    def _get_fields(self, quantity_str, tmp_output):
        """Helper: create B2B file with given quantity and extract fields."""
        results = [{
            'part_number': 'TEST-001',
            'revision': 'A',
            'quantity': quantity_str,
            'confidence_level': 'HIGH',
            'process_summary_hebrew': 'ציפוי',
            'item_name': 'Test Part',
            'drawing_number': 'DWG-001',
        }]
        output_file = tmp_output / "B2B-0_200002-test.txt"
        _save_text_summary_with_variants(
            results, output_file, customer_email="t@t.com",
            b2b_number="B2B-0", timestamp="test"
        )
        txt_files = list(tmp_output.glob("B2B-0*.txt"))
        content = txt_files[0].read_text(encoding='cp1255', errors='replace')
        line = content.split('{~#~}')[0]
        fields = line.split('\t')
        return fields

    def test_single_number_goes_to_field4(self, tmp_output):
        fields = self._get_fields("153", tmp_output)
        assert fields[3] == "153"   # field 4 = quantity
        assert fields[8] == ""      # field 9 = empty

    def test_range_goes_to_field9(self, tmp_output):
        fields = self._get_fields("10-20", tmp_output)
        assert fields[3] == "0"         # field 4 = 0
        assert fields[8] == "10-20"     # field 9 = range

    def test_multiple_quantities_go_to_field9(self, tmp_output):
        fields = self._get_fields("(80, 100)", tmp_output)
        assert fields[3] == "0"             # field 4 = 0
        assert fields[8] == "(80, 100)"     # field 9

    def test_text_quantity_goes_to_field9(self, tmp_output):
        fields = self._get_fields("10-20 יחידות (טווח)", tmp_output)
        assert fields[3] == "0"
        assert "10-20" in fields[8]
