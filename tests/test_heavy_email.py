"""Tests for the heavy email (max_files_per_email) feature."""
import zipfile
from pathlib import Path

import pytest

from automation_runner import _count_drawing_files, _DRAWING_EXTS


class TestCountDrawingFiles:
    """Tests for _count_drawing_files helper."""

    def test_empty_dir(self, tmp_path):
        assert _count_drawing_files(tmp_path) == 0

    def test_only_non_drawing_files(self, tmp_path):
        (tmp_path / "readme.txt").write_text("hello")
        (tmp_path / "data.xlsx").write_bytes(b"\x00")
        (tmp_path / "email.txt").write_text("email body")
        assert _count_drawing_files(tmp_path) == 0

    def test_pdfs_and_images(self, tmp_path):
        (tmp_path / "drawing1.pdf").write_bytes(b"%PDF")
        (tmp_path / "drawing2.tif").write_bytes(b"\x00")
        (tmp_path / "photo.jpg").write_bytes(b"\xFF\xD8")
        (tmp_path / "notes.txt").write_text("notes")
        assert _count_drawing_files(tmp_path) == 3

    def test_counts_all_drawing_extensions(self, tmp_path):
        for i, ext in enumerate(_DRAWING_EXTS):
            (tmp_path / f"file{i}{ext}").write_bytes(b"\x00")
        assert _count_drawing_files(tmp_path) == len(_DRAWING_EXTS)

    def test_zip_contents_counted(self, tmp_path):
        # Create a ZIP with 5 PDFs inside
        zip_path = tmp_path / "drawings.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            for i in range(5):
                zf.writestr(f"drawing_{i}.pdf", b"%PDF dummy")
            zf.writestr("readme.txt", "not a drawing")
        # Plus 2 loose PDFs
        (tmp_path / "loose1.pdf").write_bytes(b"%PDF")
        (tmp_path / "loose2.png").write_bytes(b"\x89PNG")
        assert _count_drawing_files(tmp_path) == 7  # 5 in zip + 2 loose

    def test_zip_with_no_drawings(self, tmp_path):
        zip_path = tmp_path / "docs.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("readme.txt", "text")
            zf.writestr("data.csv", "1,2,3")
        assert _count_drawing_files(tmp_path) == 0

    def test_corrupt_zip_ignored(self, tmp_path):
        (tmp_path / "bad.zip").write_bytes(b"not a real zip")
        (tmp_path / "good.pdf").write_bytes(b"%PDF")
        assert _count_drawing_files(tmp_path) == 1

    def test_nested_subdirectory(self, tmp_path):
        sub = tmp_path / "attachments"
        sub.mkdir()
        (sub / "a.pdf").write_bytes(b"%PDF")
        (sub / "b.tiff").write_bytes(b"\x00")
        (tmp_path / "c.jpg").write_bytes(b"\xFF\xD8")
        assert _count_drawing_files(tmp_path) == 3

    def test_zip_with_nested_folders(self, tmp_path):
        zip_path = tmp_path / "nested.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("folder/subfolder/drawing.pdf", b"%PDF")
            zf.writestr("folder/", "")  # directory entry
            zf.writestr("other.jpg", b"\xFF")
        assert _count_drawing_files(tmp_path) == 2

    def test_threshold_boundary(self, tmp_path):
        """Simulate exactly at and over the threshold."""
        for i in range(15):
            (tmp_path / f"file{i}.pdf").write_bytes(b"%PDF")
        count = _count_drawing_files(tmp_path)
        assert count == 15
        # At threshold (<=) should NOT be heavy
        threshold = 15
        assert not (count > threshold)
        # One more → over threshold
        (tmp_path / "extra.pdf").write_bytes(b"%PDF")
        assert _count_drawing_files(tmp_path) > threshold
