"""Tests for classifier.py pre-API guards (no Azure API needed)."""
import pytest
from pathlib import Path
from unittest.mock import MagicMock
from src.services.file.classifier import classify_file_type


@pytest.fixture
def mock_client():
    """Mock AzureOpenAI client — never called in these tests."""
    return MagicMock()


class TestExtensionClassification:
    """Files classified purely by extension — no API call."""

    def test_step_file_is_3d_model(self, mock_client, tmp_path):
        f = tmp_path / "part.step"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_MODEL"

    def test_stp_file_is_3d_model(self, mock_client, tmp_path):
        f = tmp_path / "assembly.stp"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_MODEL"

    def test_iges_file_is_3d_model(self, mock_client, tmp_path):
        f = tmp_path / "part.iges"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_MODEL"

    def test_x_t_file_is_3d_model(self, mock_client, tmp_path):
        f = tmp_path / "body.x_t"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_MODEL"

    def test_zip_file_is_archive(self, mock_client, tmp_path):
        f = tmp_path / "files.zip"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "ARCHIVE"

    def test_jpg_file_is_3d_image(self, mock_client, tmp_path):
        f = tmp_path / "photo.jpg"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_IMAGE"

    def test_png_file_is_3d_image(self, mock_client, tmp_path):
        f = tmp_path / "render.png"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_IMAGE"

    def test_unsupported_ext_is_other(self, mock_client, tmp_path):
        f = tmp_path / "readme.docx"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "OTHER"

    def test_stl_file_is_3d_model(self, mock_client, tmp_path):
        f = tmp_path / "print.stl"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_MODEL"


class TestFilenameClassification:
    """Files classified by filename patterns — no API call."""

    def test_pl_prefix_is_parts_list(self, mock_client, tmp_path):
        f = tmp_path / "PL_12345.pdf"
        f.write_bytes(b"%PDF-1.4 dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "PARTS_LIST"

    def test_pl_with_number_is_parts_list(self, mock_client, tmp_path):
        f = tmp_path / "PL1093Y815.pdf"
        f.write_bytes(b"%PDF-1.4 dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "PARTS_LIST"

    def test_partlist_is_parts_list(self, mock_client, tmp_path):
        f = tmp_path / "partlist_assembly.pdf"
        f.write_bytes(b"%PDF-1.4 dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "PARTS_LIST"

    def test_bom_is_parts_list(self, mock_client, tmp_path):
        f = tmp_path / "BOM_system.pdf"
        f.write_bytes(b"%PDF-1.4 dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "PARTS_LIST"

    def test_pl_with_underscore_is_parts_list(self, mock_client, tmp_path):
        """Regression: PL_TL-4341 was missed because _ is a word char."""
        f = tmp_path / "PL_TL-4341.pdf"
        f.write_bytes(b"%PDF-1.4 dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "PARTS_LIST"

    def test_number_pl_suffix_is_parts_list(self, mock_client, tmp_path):
        f = tmp_path / "778PL.pdf"
        f.write_bytes(b"%PDF-1.4 dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "PARTS_LIST"

    def test_model_keyword_is_3d_model(self, mock_client, tmp_path):
        f = tmp_path / "housing_model.pdf"
        f.write_bytes(b"%PDF-1.4 dummy")
        result = classify_file_type(str(f), mock_client)
        assert result[0] == "3D_MODEL"


class TestReturnFormat:
    """Verify classify_file_type returns correct tuple structure."""

    def test_returns_6_tuple(self, mock_client, tmp_path):
        f = tmp_path / "test.step"
        f.write_bytes(b"dummy")
        result = classify_file_type(str(f), mock_client)
        assert len(result) == 6

    def test_tuple_types(self, mock_client, tmp_path):
        f = tmp_path / "test.step"
        f.write_bytes(b"dummy")
        file_type, desc, quote, order, inp_tok, out_tok = classify_file_type(str(f), mock_client)
        assert isinstance(file_type, str)
        assert isinstance(desc, str)
        assert isinstance(inp_tok, int)
        assert isinstance(out_tok, int)

    def test_no_api_calls_for_extension(self, mock_client, tmp_path):
        """Ensure mock client is never called for extension-based classification."""
        f = tmp_path / "part.step"
        f.write_bytes(b"dummy")
        classify_file_type(str(f), mock_client)
        mock_client.chat.completions.create.assert_not_called()
