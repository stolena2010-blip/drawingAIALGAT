"""Runtime model configuration and Azure OpenAI helpers."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Dict, Optional

from openai import AzureOpenAI, OpenAI


def _normalize_azure_endpoint(endpoint: str) -> str:
    value = (endpoint or "").strip()
    if value.endswith("/openai/v1/"):
        value = value[:-11]
    elif value.endswith("/openai/v1"):
        value = value[:-10]
    return value.rstrip("/")


def _safe_float(value: str, default: float) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _safe_int(value: str, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def _try_float(value: str) -> Optional[float]:
    try:
        return float(value)
    except Exception:
        return None


def _normalize_model_env_key(model_name: str) -> str:
    key = re.sub(r"[^A-Za-z0-9]+", "_", (model_name or "").strip().upper())
    key = re.sub(r"_+", "_", key).strip("_")
    return key


@dataclass(frozen=True)
class ModelRuntimeConfig:
    endpoint: str
    api_key: str
    api_version: str
    deployment: str
    input_price_per_1m: float
    output_price_per_1m: float

    @classmethod
    def from_env(cls) -> "ModelRuntimeConfig":
        return cls(
            endpoint=_normalize_azure_endpoint(os.getenv("AZURE_OPENAI_ENDPOINT", "")),
            api_key=os.getenv("AZURE_OPENAI_API_KEY", ""),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview"),
            deployment=os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o"),
            input_price_per_1m=_safe_float(os.getenv("AZURE_MODEL_INPUT_PRICE_PER_1M", "2.50"), 2.50),
            output_price_per_1m=_safe_float(os.getenv("AZURE_MODEL_OUTPUT_PRICE_PER_1M", "10.00"), 10.00),
        )

    def get_stage_model(self, stage_num: int) -> str:
        return (os.getenv(f"STAGE_{stage_num}_MODEL", "") or "").strip() or self.deployment

    def _get_model_input_price(self, model_name: str) -> float:
        model_key = _normalize_model_env_key(model_name)
        if not model_key:
            return self.input_price_per_1m

        direct_key = f"MODEL_{model_key}_INPUT_PRICE"
        value = os.getenv(direct_key, "")
        if value.strip():
            return _safe_float(value, self.input_price_per_1m)
        return self.input_price_per_1m

    def _get_model_output_price(self, model_name: str) -> float:
        model_key = _normalize_model_env_key(model_name)
        if not model_key:
            return self.output_price_per_1m

        direct_key = f"MODEL_{model_key}_OUTPUT_PRICE"
        value = os.getenv(direct_key, "")
        if value.strip():
            return _safe_float(value, self.output_price_per_1m)
        return self.output_price_per_1m

    def get_stage_input_price(self, stage_num: int) -> float:
        value = os.getenv(f"STAGE_{stage_num}_INPUT_PRICE", "")
        if value.strip():
            parsed = _try_float(value)
            if parsed is not None:
                return parsed
        stage_model = self.get_stage_model(stage_num)
        return self._get_model_input_price(stage_model)

    def get_stage_output_price(self, stage_num: int) -> float:
        value = os.getenv(f"STAGE_{stage_num}_OUTPUT_PRICE", "")
        if value.strip():
            parsed = _try_float(value)
            if parsed is not None:
                return parsed
        stage_model = self.get_stage_model(stage_num)
        return self._get_model_output_price(stage_model)

    def get_stage_temperature(self, stage_num: int, default: float) -> float:
        value = os.getenv(f"STAGE_{stage_num}_TEMPERATURE", "")
        if value.strip():
            return _safe_float(value, default)
        return default

    def get_stage_max_tokens(self, stage_num: int, default: int) -> int:
        value = os.getenv(f"STAGE_{stage_num}_MAX_TOKENS", "")
        if value.strip():
            return max(1, _safe_int(value, default))
        return default

    # ── Per-model endpoint / key / version (for models on a different Azure resource) ──

    def get_model_endpoint(self, model_name: str) -> str:
        """Return model-specific endpoint, or the global default."""
        model_key = _normalize_model_env_key(model_name)
        if model_key:
            value = os.getenv(f"MODEL_{model_key}_ENDPOINT", "").strip()
            if value:
                return _normalize_azure_endpoint(value)
        return self.endpoint

    def get_model_api_key(self, model_name: str) -> str:
        """Return model-specific API key, or the global default."""
        model_key = _normalize_model_env_key(model_name)
        if model_key:
            value = os.getenv(f"MODEL_{model_key}_API_KEY", "").strip()
            if value:
                return value
        return self.api_key

    def get_model_api_version(self, model_name: str) -> str:
        """Return model-specific API version, or the global default."""
        model_key = _normalize_model_env_key(model_name)
        if model_key:
            value = os.getenv(f"MODEL_{model_key}_API_VERSION", "").strip()
            if value:
                return value
        return self.api_version

    def get_model_deployment(self, model_name: str) -> str:
        """Return Azure deployment name for *model_name*.

        Checks ``MODEL_<KEY>_DEPLOYMENT`` env var first; if not set, falls
        back to using *model_name* itself as the deployment name (which is
        the common case when model name == deployment name).
        """
        model_key = _normalize_model_env_key(model_name)
        if model_key:
            value = os.getenv(f"MODEL_{model_key}_DEPLOYMENT", "").strip()
            if value:
                return value
        return model_name or self.deployment

    def is_model_openai_compat(self, model_name: str) -> bool:
        """Check if model requires OpenAI-compatible client (base_url) instead of AzureOpenAI."""
        model_key = _normalize_model_env_key(model_name)
        if model_key:
            value = os.getenv(f"MODEL_{model_key}_USE_OPENAI_CLIENT", "").strip().lower()
            if value in {"1", "true", "yes"}:
                return True
        return False

    def is_model_reasoning(self, model_name: str) -> bool:
        """Check if model is a reasoning model (o-series or explicitly flagged)."""
        if (model_name or "").strip().lower().startswith("o"):
            return True
        model_key = _normalize_model_env_key(model_name)
        if model_key:
            value = os.getenv(f"MODEL_{model_key}_IS_REASONING", "").strip().lower()
            if value in {"1", "true", "yes"}:
                return True
        return False


# ── Client cache (one client per unique endpoint + key combination) ──
_client_cache: Dict[str, object] = {}   # AzureOpenAI or OpenAI


def build_azure_client(
    config: Optional[ModelRuntimeConfig] = None,
    model_name: Optional[str] = None,
):
    """Build (or return cached) AzureOpenAI *or* OpenAI client.

    If *model_name* is given and has its own endpoint/key configured via
    ``MODEL_<NAME>_ENDPOINT`` / ``MODEL_<NAME>_API_KEY`` env vars, a separate
    client pointing at that resource is created and cached.

    When ``MODEL_<NAME>_USE_OPENAI_CLIENT=true``, an ``OpenAI`` client with
    ``base_url`` is returned (for Azure resources that expose the
    OpenAI-compatible ``/openai/v1`` endpoint, e.g. gpt-5.4).
    """
    runtime = config or ModelRuntimeConfig.from_env()

    if model_name:
        endpoint = runtime.get_model_endpoint(model_name)
        api_key = runtime.get_model_api_key(model_name)
        api_version = runtime.get_model_api_version(model_name)
        use_openai = runtime.is_model_openai_compat(model_name)
    else:
        endpoint = runtime.endpoint
        api_key = runtime.api_key
        api_version = runtime.api_version
        use_openai = False

    client_type = "openai" if use_openai else "azure"
    cache_key = f"{client_type}|{endpoint}|{api_key[:12]}|{api_version}"
    if cache_key in _client_cache:
        return _client_cache[cache_key]

    if use_openai:
        # OpenAI-compatible endpoint (e.g. https://host.openai.azure.com/openai/v1)
        base_url = endpoint.rstrip("/")
        if not base_url.endswith("/openai/v1"):
            base_url += "/openai/v1"
        client = OpenAI(
            base_url=base_url,
            api_key=api_key,
            timeout=120.0,
        )
    else:
        client = AzureOpenAI(
            azure_endpoint=endpoint,
            api_key=api_key,
            api_version=api_version,
            timeout=120.0,
        )

    _client_cache[cache_key] = client
    return client


def calculate_token_cost(
    input_tokens: int,
    output_tokens: int,
    config: Optional[ModelRuntimeConfig] = None,
) -> float:
    runtime = config or ModelRuntimeConfig.from_env()
    input_cost = (input_tokens / 1_000_000) * runtime.input_price_per_1m
    output_cost = (output_tokens / 1_000_000) * runtime.output_price_per_1m
    return input_cost + output_cost
