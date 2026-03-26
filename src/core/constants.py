"""
Core constants, config and helpers for DrawingAI Pro.
=====================================================

Centralised definitions extracted from customer_extractor_v3_dual.py
so that every module can ``from src.core.constants import …`` without
pulling in the full extractor.
"""

import os
import sys

from src.services.ai import ModelRuntimeConfig

# ---------------------------------------------------------------------------
# GUI stop / skip hooks (set by the GUI layer at runtime)
# ---------------------------------------------------------------------------
_gui_should_stop = None
_gui_should_skip = None

# ---------------------------------------------------------------------------
# Debug / feature flags
# ---------------------------------------------------------------------------
DEBUG_ENABLED = os.getenv("AI_DRAW_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}
IAI_TOP_RED_FALLBACK_ENABLED = os.getenv("IAI_TOP_RED_FALLBACK_ENABLED", "true").strip().lower() in {"1", "true", "yes", "on"}


def debug_print(message: str) -> None:
    """Print a debug message when AI_DRAW_DEBUG is enabled."""
    if DEBUG_ENABLED:
        print(message)


# ---------------------------------------------------------------------------
# Azure OpenAI runtime configuration
# ---------------------------------------------------------------------------
MODEL_RUNTIME = ModelRuntimeConfig.from_env()
AZURE_DEPLOYMENT = MODEL_RUNTIME.deployment
MODEL_INPUT_PRICE_PER_1M = MODEL_RUNTIME.input_price_per_1m
MODEL_OUTPUT_PRICE_PER_1M = MODEL_RUNTIME.output_price_per_1m

# ---------------------------------------------------------------------------
# Supported drawing file extensions
# ---------------------------------------------------------------------------
DRAWING_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}

# ---------------------------------------------------------------------------
# Per-stage model mapping (configurable in .env via STAGE_{N}_*)
# ---------------------------------------------------------------------------
STAGE_CLASSIFICATION = 0
STAGE_LAYOUT = 0
STAGE_ROTATION = 0
STAGE_BASIC_INFO = 1
STAGE_PROCESSES = 2
STAGE_NOTES = 3
STAGE_AREA = 4
STAGE_VALIDATION = 5
STAGE_PL = 6
STAGE_EMAIL_QUANTITIES = 7
STAGE_ORDER_ITEM_DETAILS = 8

STAGE_DISPLAY_NAMES = {
    STAGE_CLASSIFICATION: "Stage 0 (Classification/Layout/Rotation)",
    STAGE_BASIC_INFO: "Stage 1 (Basic Info)",
    STAGE_PROCESSES: "Stage 2 (Processes)",
    STAGE_NOTES: "Stage 3 (Notes)",
    STAGE_AREA: "Stage 4 (Area)",
    STAGE_VALIDATION: "Stage 5 (Validation)",
    STAGE_PL: "Stage 6 (PL)",
    STAGE_EMAIL_QUANTITIES: "Stage 7 (Email Quantities)",
    STAGE_ORDER_ITEM_DETAILS: "Stage 8 (Quote/Order Item Details)",
}

# ---------------------------------------------------------------------------
# File-size limits (MB)
# ---------------------------------------------------------------------------
MAX_FILE_SIZE_MB = 100   # Skip files larger than this
WARN_FILE_SIZE_MB = 50   # Warn about files larger than this

# ---------------------------------------------------------------------------
# Image resolution limits (pixels)
# ---------------------------------------------------------------------------
MAX_IMAGE_DIMENSION = 4096     # Maximum width or height in pixels
TARGET_IMAGE_DIMENSION = 2048  # Target dimension for downsampling high-res images
WARN_IMAGE_DIMENSION = 3000    # Warn about images larger than this

# ---------------------------------------------------------------------------
# Structured JSON output format for OpenAI calls
# ---------------------------------------------------------------------------
RESPONSE_FORMAT = {"type": "json_object"}
