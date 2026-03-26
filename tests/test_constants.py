"""Unit tests for src/core/constants.py"""

from src.core.constants import (
    STAGE_BASIC_INFO, STAGE_PROCESSES, STAGE_NOTES, STAGE_AREA,
    STAGE_DISPLAY_NAMES, DRAWING_EXTS, MAX_FILE_SIZE_MB,
    MAX_IMAGE_DIMENSION, RESPONSE_FORMAT, debug_print,
)


def test_stage_numbers_are_sequential():
    assert STAGE_BASIC_INFO == 1
    assert STAGE_PROCESSES == 2
    assert STAGE_NOTES == 3
    assert STAGE_AREA == 4


def test_stage_display_names_has_all_stages():
    assert STAGE_BASIC_INFO in STAGE_DISPLAY_NAMES
    assert STAGE_PROCESSES in STAGE_DISPLAY_NAMES
    assert STAGE_NOTES in STAGE_DISPLAY_NAMES
    assert STAGE_AREA in STAGE_DISPLAY_NAMES


def test_drawing_extensions():
    assert ".pdf" in DRAWING_EXTS
    assert ".png" in DRAWING_EXTS
    assert ".jpg" in DRAWING_EXTS
    assert ".tif" in DRAWING_EXTS
    assert ".docx" not in DRAWING_EXTS


def test_file_size_limits():
    assert MAX_FILE_SIZE_MB > 0
    assert MAX_IMAGE_DIMENSION > 0


def test_response_format():
    assert RESPONSE_FORMAT == {"type": "json_object"}


def test_debug_print_does_not_crash():
    # Should not raise even if debug is off
    debug_print("test message")
