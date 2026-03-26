"""Tests for validate_notes_before_stage5 — pure logic, no API."""
import pytest
from customer_extractor_v3_dual import validate_notes_before_stage5


class TestValidateNotes:
    def test_none_input(self):
        result, report = validate_notes_before_stage5(None)
        assert result is None
        assert not report["is_valid"]

    def test_empty_string(self):
        result, report = validate_notes_before_stage5("")
        assert result is None
        assert not report["is_valid"]

    def test_short_text(self):
        result, report = validate_notes_before_stage5("hi")
        # Short text (< 20 chars) → valid but with "Very short" warning
        assert report["char_count"] == 2
        assert any("Very short" in w for w in report["warnings"])

    def test_valid_english(self):
        text = "Apply passivation per AMS 2700 after machining. Mask holes."
        result, report = validate_notes_before_stage5(text)
        assert result is not None
        assert report["has_numbers"]
        assert report["is_valid"]

    def test_valid_hebrew(self):
        text = "ציפוי קשיח לפי מפרט AMS 2700. מיסוך חורים בקוטר מתחת ל-5 מילימטר."
        result, report = validate_notes_before_stage5(text)
        assert result is not None
        assert report["has_hebrew"]
        assert report["is_valid"]

    def test_encoding_ok(self):
        text = "Normal text with UTF-8 chars: testing encoding check here"
        result, report = validate_notes_before_stage5(text)
        assert report["encoding_ok"]

    def test_long_text_warning(self):
        text = "A" * 6000
        result, report = validate_notes_before_stage5(text)
        assert any("Very long" in w for w in report["warnings"])

    def test_integer_input(self):
        result, report = validate_notes_before_stage5(123)
        assert result is None
        assert "Invalid type" in report["issues"][0]

    def test_whitespace_only(self):
        result, report = validate_notes_before_stage5("   \n\t  ")
        assert result is None
        assert not report["is_valid"]

    def test_numbers_detected(self):
        text = "Surface roughness Ra 0.8 as specified in drawing"
        result, report = validate_notes_before_stage5(text)
        assert report["has_numbers"]

    def test_control_chars_warning(self):
        text = "Normal text with\x01control\x02chars inside the notes area"
        result, report = validate_notes_before_stage5(text)
        assert any("problematic" in w for w in report["warnings"])
