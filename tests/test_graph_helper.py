"""Unit tests for src/services/email/graph_helper.py (pure helpers only)."""

from src.services.email.graph_helper import (
    _sanitize_filename, _clean_html_body, _save_email_txt,
)
from pathlib import Path
import tempfile


# =========================================================================
# _sanitize_filename tests
# =========================================================================

def test_sanitize_removes_tabs():
    assert "\t" not in _sanitize_filename("file\tname.pdf")

def test_sanitize_removes_special_chars():
    result = _sanitize_filename('file<>:"/\\|?*name.pdf')
    assert "<" not in result
    assert ">" not in result
    assert ":" not in result
    assert '"' not in result

def test_sanitize_truncates_long_names():
    long_name = "a" * 300
    assert len(_sanitize_filename(long_name)) <= 180

def test_sanitize_collapses_whitespace():
    assert "  " not in _sanitize_filename("file   name   .pdf")

def test_sanitize_normal_name_unchanged():
    assert _sanitize_filename("drawing_3585410.pdf") == "drawing_3585410.pdf"


# =========================================================================
# _clean_html_body tests
# =========================================================================

def test_clean_html_strips_tags():
    result = _clean_html_body("<p>Hello <b>World</b></p>")
    assert "<p>" not in result
    assert "<b>" not in result
    assert "Hello" in result
    assert "World" in result

def test_clean_html_converts_br():
    result = _clean_html_body("Line1<br>Line2<br/>Line3")
    assert "\n" in result

def test_clean_html_decodes_entities():
    result = _clean_html_body("5 &gt; 3 &amp; 2 &lt; 4")
    assert ">" in result
    assert "&" in result
    assert "<" in result

def test_clean_html_removes_signatures():
    result = _clean_html_body("Hello\nBest regards,\nJohn Smith\nCEO")
    assert "Best regards" not in result

def test_clean_html_empty():
    assert _clean_html_body("") == ""


# =========================================================================
# _save_email_txt tests
# =========================================================================

def test_save_email_txt_creates_file():
    with tempfile.TemporaryDirectory() as tmpdir:
        msg_dir = Path(tmpdir)
        _save_email_txt(msg_dir, "sender@test.com", "Test Subject", "2026-02-18T10:00:00Z", "Body text")
        email_file = msg_dir / "email.txt"
        assert email_file.exists()
        content = email_file.read_text(encoding="utf-8")
        assert "sender@test.com" in content
        assert "Test Subject" in content
        assert "Body text" in content

def test_save_email_txt_empty_body():
    with tempfile.TemporaryDirectory() as tmpdir:
        msg_dir = Path(tmpdir)
        _save_email_txt(msg_dir, "a@b.com", "Sub", "", "")
        content = (msg_dir / "email.txt").read_text(encoding="utf-8")
        assert "(ללא תוכן)" in content
