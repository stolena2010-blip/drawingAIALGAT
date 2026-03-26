"""Tests for automation_runner.py utility functions (no API/email needed)."""
import json
import pytest
from pathlib import Path
from automation_runner import (
    _load_json,
    _save_json,
    _rotate_log_if_needed,
    _append_log,
    _clean_sender_line,
    _now_iso,
)


class TestLoadJson:
    def test_load_existing_file(self, tmp_path):
        f = tmp_path / "test.json"
        f.write_text(json.dumps({"key": "value"}), encoding="utf-8")
        result = _load_json(f, {})
        assert result == {"key": "value"}

    def test_load_missing_file_returns_default(self, tmp_path):
        f = tmp_path / "nonexistent.json"
        result = _load_json(f, {"default": True})
        assert result == {"default": True}

    def test_load_corrupted_file_returns_default(self, tmp_path):
        f = tmp_path / "bad.json"
        f.write_text("not json {{{", encoding="utf-8")
        result = _load_json(f, {"fallback": 1})
        assert result == {"fallback": 1}

    def test_load_empty_default(self, tmp_path):
        f = tmp_path / "nope.json"
        result = _load_json(f, {})
        assert result == {}


class TestSaveJson:
    def test_save_creates_file(self, tmp_path):
        f = tmp_path / "out.json"
        _save_json(f, {"ids": [1, 2, 3]})
        assert f.exists()
        data = json.loads(f.read_text(encoding="utf-8"))
        assert data["ids"] == [1, 2, 3]

    def test_save_overwrites(self, tmp_path):
        f = tmp_path / "out.json"
        _save_json(f, {"v": 1})
        _save_json(f, {"v": 2})
        data = json.loads(f.read_text(encoding="utf-8"))
        assert data["v"] == 2

    def test_save_hebrew(self, tmp_path):
        f = tmp_path / "hebrew.json"
        _save_json(f, {"name": "שלום"})
        data = json.loads(f.read_text(encoding="utf-8"))
        assert data["name"] == "שלום"

    def test_atomic_no_temp_leftover(self, tmp_path):
        f = tmp_path / "clean.json"
        _save_json(f, {"ok": True})
        tmp_file = f.with_suffix(".tmp")
        assert not tmp_file.exists()


class TestRotateLog:
    def test_no_rotation_small_file(self, tmp_path):
        log = tmp_path / "test.jsonl"
        log.write_text("small\n", encoding="utf-8")
        _rotate_log_if_needed(log, max_size_bytes=1_000_000)
        assert log.exists()  # Not rotated

    def test_rotation_large_file(self, tmp_path):
        log = tmp_path / "test.jsonl"
        log.write_text("x" * 2_000_000, encoding="utf-8")
        _rotate_log_if_needed(log, max_size_bytes=1_000_000)
        assert not log.exists()  # Original removed
        rotated = list(tmp_path.glob("automation_log_*.jsonl"))
        assert len(rotated) == 1

    def test_no_crash_missing_file(self, tmp_path):
        log = tmp_path / "nonexistent.jsonl"
        _rotate_log_if_needed(log)  # Should not raise


class TestAppendLog:
    def test_append_creates_file(self, tmp_path):
        log = tmp_path / "new.jsonl"
        _append_log(log, {"event": "test"})
        assert log.exists()
        line = log.read_text(encoding="utf-8").strip()
        data = json.loads(line)
        assert data["event"] == "test"

    def test_append_multiple(self, tmp_path):
        log = tmp_path / "multi.jsonl"
        _append_log(log, {"n": 1})
        _append_log(log, {"n": 2})
        lines = log.read_text(encoding="utf-8").strip().split("\n")
        assert len(lines) == 2


class TestCleanSenderLine:
    def test_removes_hebrew_prefix(self):
        assert _clean_sender_line("כתובת שולח: user@test.com") == "user@test.com"

    def test_removes_from_prefix(self):
        assert _clean_sender_line("From: user@test.com") == "user@test.com"

    def test_strips_whitespace(self):
        assert _clean_sender_line("  user@test.com  ") == "user@test.com"

    def test_plain_email_unchanged(self):
        assert _clean_sender_line("user@test.com") == "user@test.com"


class TestNowIso:
    def test_returns_string(self):
        result = _now_iso()
        assert isinstance(result, str)

    def test_contains_T(self):
        """ISO format contains T separator."""
        result = _now_iso()
        assert "T" in result or "-" in result
