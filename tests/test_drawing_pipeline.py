"""Tests for src/services/extraction/drawing_pipeline.py

Tests cover the pure helper functions and constants that do not require
a live Azure OpenAI client or OCR engine.
"""
import os
import time
import pytest
from src.services.extraction.drawing_pipeline import (
    DRAWING_EXTS,
    MAX_IMAGE_DIMENSION,
    WARN_IMAGE_DIMENSION,
    _run_with_timeout,
)


# ── DRAWING_EXTS constant ─────────────────────────────────────────────

class TestDrawingExts:

    def test_pdf_supported(self):
        assert ".pdf" in DRAWING_EXTS

    def test_png_supported(self):
        assert ".png" in DRAWING_EXTS

    def test_jpg_supported(self):
        assert ".jpg" in DRAWING_EXTS

    def test_jpeg_supported(self):
        assert ".jpeg" in DRAWING_EXTS

    def test_tif_supported(self):
        assert ".tif" in DRAWING_EXTS

    def test_tiff_supported(self):
        assert ".tiff" in DRAWING_EXTS

    def test_txt_not_supported(self):
        assert ".txt" not in DRAWING_EXTS

    def test_docx_not_supported(self):
        assert ".docx" not in DRAWING_EXTS

    def test_all_lowercase(self):
        # All extensions must be lower-case to match Path().suffix.lower()
        for ext in DRAWING_EXTS:
            assert ext == ext.lower(), f"Extension '{ext}' is not lower-case"


# ── Image dimension constants ─────────────────────────────────────────

class TestImageDimensionConstants:

    def test_max_dimension_positive(self):
        assert MAX_IMAGE_DIMENSION > 0

    def test_warn_below_max(self):
        assert WARN_IMAGE_DIMENSION < MAX_IMAGE_DIMENSION

    def test_reasonable_values(self):
        # Sanity: not absurdly small or impossibly large
        assert 1000 < MAX_IMAGE_DIMENSION <= 16384


# ── _run_with_timeout ─────────────────────────────────────────────────

class TestRunWithTimeout:

    def test_fast_function_returns_result(self):
        result = _run_with_timeout(lambda: 42, timeout=5)
        assert result == 42

    def test_function_with_args(self):
        result = _run_with_timeout(lambda a, b: a + b, args=(3, 4), timeout=5)
        assert result == 7

    def test_function_with_kwargs(self):
        def greet(name="world"):
            return f"hello {name}"

        result = _run_with_timeout(greet, kwargs={"name": "test"}, timeout=5)
        assert result == "hello test"

    def test_timeout_returns_none(self):
        def slow_fn():
            time.sleep(10)
            return "done"

        result = _run_with_timeout(slow_fn, timeout=1, stage_name="slow")
        assert result is None

    def test_exception_inside_function_returns_none(self):
        def boom():
            raise ValueError("unexpected!")

        result = _run_with_timeout(boom, timeout=5, stage_name="boom")
        assert result is None

    def test_none_kwargs_default(self):
        # kwargs=None should be handled gracefully (defaulted to {})
        result = _run_with_timeout(lambda: "ok", kwargs=None, timeout=5)
        assert result == "ok"

    def test_returns_none_value_from_function(self):
        # A function that legitimately returns None should not be confused with timeout
        result = _run_with_timeout(lambda: None, timeout=5)
        assert result is None

    def test_stage_timeout_env_override(self, monkeypatch):
        # STAGE_HARD_TIMEOUT env var is read at module import time,
        # but the per-call timeout= parameter always takes precedence.
        monkeypatch.setenv("STAGE_HARD_TIMEOUT_SECONDS", "999")
        result = _run_with_timeout(lambda: "env ok", timeout=5)
        assert result == "env ok"
