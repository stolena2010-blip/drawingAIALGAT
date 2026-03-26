"""Unit tests for src/services/ai/vision_api.py (no real API calls)."""

import pytest
from unittest.mock import patch, MagicMock
from src.services.ai.vision_api import (
    _resolve_stage_call_config,
    _calculate_stage_cost,
)


def test_resolve_stage_call_config_defaults():
    """When stage_num is None, should return defaults."""
    deployment, max_tokens, temperature = _resolve_stage_call_config(
        stage_num=None, default_max_tokens=4000, default_temperature=0.0
    )
    assert max_tokens == 4000
    assert temperature == 0.0
    assert isinstance(deployment, str)


def test_resolve_stage_call_config_with_stage():
    """When stage_num is provided, should return config for that stage."""
    deployment, max_tokens, temperature = _resolve_stage_call_config(
        stage_num=1, default_max_tokens=4000, default_temperature=0.0
    )
    assert isinstance(deployment, str)
    assert isinstance(max_tokens, int)
    assert isinstance(temperature, float)


def test_calculate_stage_cost():
    """Cost should be positive for non-zero tokens."""
    cost = _calculate_stage_cost(input_tokens=1000, output_tokens=500, stage_num=1)
    assert isinstance(cost, float)
    assert cost >= 0


def test_calculate_stage_cost_zero():
    """Zero tokens = zero cost."""
    cost = _calculate_stage_cost(input_tokens=0, output_tokens=0, stage_num=1)
    assert cost == 0.0
