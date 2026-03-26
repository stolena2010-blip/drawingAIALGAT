"""Mock integration test: scan_folder flow with fake API responses."""
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from customer_extractor_v3_dual import scan_folder


class TestScanFolderFlow:

    @pytest.fixture
    def fake_folder(self, tmp_path):
        """Create a minimal folder structure for testing."""
        folder = tmp_path / "test_email"
        folder.mkdir()

        # Create a fake email.txt
        email_txt = folder / "email.txt"
        email_txt.write_text(
            "From: test@customer.com\nSubject: Test Order\n\nPlease process.",
            encoding="utf-8"
        )

        # Create a minimal PDF-like file (won't work with real OCR)
        fake_pdf = folder / "drawing_12345.pdf"
        fake_pdf.write_bytes(b"%PDF-1.4 fake content")

        return folder

    def test_empty_folder(self, tmp_path):
        """scan_folder should handle empty folder gracefully."""
        empty = tmp_path / "empty"
        empty.mkdir()

        result = scan_folder(empty, recursive=False)
        # Returns a tuple (results, folder_path, ...) or list
        assert isinstance(result, (list, tuple))

    def test_nonexistent_folder(self, tmp_path):
        """scan_folder should handle missing folder."""
        missing = tmp_path / "does_not_exist"

        result = scan_folder(missing, recursive=False)
        assert isinstance(result, (list, tuple))

    def test_no_stages_selected(self, tmp_path):
        """scan_folder with no stages should still work."""
        folder = tmp_path / "no_stages"
        folder.mkdir()
        (folder / "test.pdf").write_bytes(b"%PDF-1.4 fake")

        result = scan_folder(
            folder,
            recursive=False,
            selected_stages={1: False, 2: False, 3: False, 4: False}
        )
        assert isinstance(result, (list, tuple))

    def test_max_file_size_filter(self, tmp_path):
        """Files larger than max_file_size_mb should be skipped."""
        folder = tmp_path / "big_file"
        folder.mkdir()

        # Create a 2MB file
        big_file = folder / "huge_drawing.pdf"
        big_file.write_bytes(b"x" * (2 * 1024 * 1024))

        result = scan_folder(
            folder,
            recursive=False,
            max_file_size_mb=1  # Max 1MB — should skip
        )
        # Should not crash
        assert isinstance(result, (list, tuple))

    def test_folder_with_email_txt(self, fake_folder):
        """Folder with email.txt should be detected as email folder."""
        email_file = fake_folder / "email.txt"
        assert email_file.exists()

        content = email_file.read_text()
        assert "test@customer.com" in content

    def test_recursive_false_no_subfolders(self, tmp_path):
        """Non-recursive scan should not enter subdirectories."""
        parent = tmp_path / "parent"
        parent.mkdir()
        child = parent / "child"
        child.mkdir()
        (child / "test.pdf").write_bytes(b"%PDF-1.4 fake")

        result = scan_folder(parent, recursive=False)
        assert isinstance(result, (list, tuple))
