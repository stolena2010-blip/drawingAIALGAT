"""
Azure OpenAI Vision API helpers for DrawingAI Pro.
====================================================

Client construction, per-stage configuration resolution, cost calculation,
token-compatibility wrapper, and retry logic — extracted from
customer_extractor_v3_dual.py.
"""

import re
from typing import Optional, Tuple

from openai import AzureOpenAI

from src.core.constants import (
    AZURE_DEPLOYMENT,
    MODEL_RUNTIME,
    RESPONSE_FORMAT,
    STAGE_DISPLAY_NAMES,
    debug_print,
)
from src.services.ai import build_azure_client


# =========================================================================
# Client helpers
# =========================================================================

def _build_client():
    """Initialize Azure OpenAI (or OpenAI-compatible) client."""
    return build_azure_client(MODEL_RUNTIME)


def _get_client_for_model(model_name: str, fallback_client=None):
    """Return the right client for *model_name*.

    If the model has its own endpoint configured (``MODEL_<KEY>_ENDPOINT``),
    or requires an OpenAI-compatible client (``MODEL_<KEY>_USE_OPENAI_CLIENT``),
    a dedicated client is built (or returned from cache).  Otherwise the
    *fallback_client* (or the default client) is returned.
    """
    if model_name:
        model_endpoint = MODEL_RUNTIME.get_model_endpoint(model_name)
        use_openai = MODEL_RUNTIME.is_model_openai_compat(model_name)
        if model_endpoint != MODEL_RUNTIME.endpoint or use_openai:
            return build_azure_client(MODEL_RUNTIME, model_name=model_name)
    return fallback_client or _build_client()


# =========================================================================
# Per-stage configuration
# =========================================================================

def _resolve_stage_call_config(
    stage_num: Optional[int],
    default_max_tokens: int,
    default_temperature: float,
) -> Tuple[str, int, float]:
    if stage_num is None:
        debug_print(
            f"[MODEL] Fallback deployment={AZURE_DEPLOYMENT} | "
            f"max_tokens={default_max_tokens} | temperature={default_temperature}"
        )
        return AZURE_DEPLOYMENT, default_max_tokens, default_temperature

    stage_model = MODEL_RUNTIME.get_stage_model(stage_num)
    stage_max_tokens = MODEL_RUNTIME.get_stage_max_tokens(stage_num, default_max_tokens)
    stage_temperature = MODEL_RUNTIME.get_stage_temperature(stage_num, default_temperature)
    stage_name = STAGE_DISPLAY_NAMES.get(stage_num, f"Stage {stage_num}")
    debug_print(
        f"[MODEL] {stage_name} -> model={stage_model} | "
        f"max_tokens={stage_max_tokens} | temperature={stage_temperature}"
    )
    return stage_model, stage_max_tokens, stage_temperature


# =========================================================================
# Cost calculation
# =========================================================================

def _calculate_stage_cost(input_tokens: int, output_tokens: int, stage_num: int) -> float:
    input_price = MODEL_RUNTIME.get_stage_input_price(stage_num)
    output_price = MODEL_RUNTIME.get_stage_output_price(stage_num)
    return (input_tokens / 1_000_000) * input_price + (output_tokens / 1_000_000) * output_price


def _log_stage_completion(response, stage_num: Optional[int], model_used: str) -> None:
    """Print per-stage model, tokens, and cost after each API call."""
    if response is None or not hasattr(response, "usage") or response.usage is None:
        return
    usage = response.usage
    in_tok = usage.prompt_tokens or 0
    out_tok = usage.completion_tokens or 0
    stage_label = STAGE_DISPLAY_NAMES.get(stage_num, f"Stage {stage_num}") if stage_num is not None else "Unknown Stage"
    if stage_num is not None:
        cost = _calculate_stage_cost(in_tok, out_tok, stage_num)
    else:
        cost = 0.0
    print(
        f"      💲 {stage_label} | model={model_used} | "
        f"tokens: {in_tok:,} in / {out_tok:,} out | cost=${cost:.4f}"
    )


# =========================================================================
# Chat completion with parameter-compatibility retries
# =========================================================================

def _chat_create_with_token_compat(client, **kwargs):
    """Create chat completion with generic compatibility retries for model-specific parameter constraints.

    If the requested *model* lives on a different Azure endpoint (detected via
    ``MODEL_<NAME>_ENDPOINT`` env var), the appropriate client is used
    automatically — callers do not need to worry about multi-endpoint routing.
    """
    # Auto-switch client if the model needs a different endpoint
    model_name = kwargs.get("model", "")
    if model_name:
        client = _get_client_for_model(model_name, fallback_client=client)
        # Resolve Azure deployment name (may differ from model name)
        deployment_name = MODEL_RUNTIME.get_model_deployment(model_name)
        if deployment_name != model_name:
            kwargs = dict(kwargs)
            kwargs["model"] = deployment_name

    request_kwargs = dict(kwargs)
    request_kwargs.setdefault("timeout", 120)
    max_attempts = 5

    for _ in range(max_attempts):
        try:
            return client.chat.completions.create(**request_kwargs)
        except Exception as error:
            error_text = str(error)
            error_str = error_text.lower()

            param_match = re.search(r"unsupported parameter:\s*'([^']+)'", error_str)
            if not param_match:
                param_match = re.search(r"['\"]param['\"]\s*:\s*['\"]([^'\"]+)['\"]", error_str)
            err_param = param_match.group(1) if param_match else None

            if (
                "max_completion_tokens" in error_str
                and "max_tokens" in request_kwargs
                and "unsupported parameter" in error_str
            ):
                request_kwargs["max_completion_tokens"] = request_kwargs.pop("max_tokens")
                debug_print("[MODEL] Retrying completion with max_completion_tokens compatibility mode")
                continue

            if err_param == "temperature" and "temperature" in request_kwargs:
                if "only the default" in error_str and "is supported" in error_str:
                    request_kwargs["temperature"] = 1
                    debug_print("[MODEL] Retrying completion with temperature=1 compatibility mode")
                else:
                    request_kwargs.pop("temperature", None)
                    debug_print("[MODEL] Retrying completion without temperature parameter")
                continue

            if err_param and err_param in request_kwargs and "unsupported" in error_str:
                request_kwargs.pop(err_param, None)
                debug_print(f"[MODEL] Retrying completion without unsupported parameter '{err_param}'")
                continue

            if "response_format" in request_kwargs and "response_format" in error_str and "unsupported" in error_str:
                request_kwargs.pop("response_format", None)
                debug_print("[MODEL] Retrying completion without response_format")
                continue

            if "seed" in request_kwargs and "seed" in error_str and "unsupported" in error_str:
                request_kwargs.pop("seed", None)
                debug_print("[MODEL] Retrying completion without seed")
                continue

            raise RuntimeError(error_text) from error

    raise RuntimeError("Chat completion failed after compatibility retries")


# =========================================================================
# Vision API with content-filter retry
# =========================================================================

def _call_vision_api_with_retry(client, messages, max_tokens=4000, temperature=0, stage_num: Optional[int] = None):
    """
    Call Azure OpenAI Vision API with automatic retry if content filter triggers.

    Args:
        client: AzureOpenAI client instance
        messages: List of message dicts for the API call
        max_tokens: Maximum tokens in response
        temperature: Temperature setting
        stage_num: Optional stage number for per-stage config

    Returns:
        API response object or None if both attempts fail

    Raises:
        Exception: If error is not content filter related
    """
    model_name, resolved_max_tokens, resolved_temperature = _resolve_stage_call_config(
        stage_num,
        max_tokens,
        temperature,
    )

    try:
        # First attempt with high-detail
        response = _chat_create_with_token_compat(
            client,
            model=model_name,
            messages=messages,
            max_tokens=resolved_max_tokens,
            temperature=resolved_temperature,
            response_format=RESPONSE_FORMAT
        )
        _log_stage_completion(response, stage_num, model_name)
        return response
    except Exception as e:
        error_str = str(e)

        if "429" in error_str or "rate limit" in error_str.lower() or "throttl" in error_str.lower():
            import time as _time
            retry_after = 10
            try:
                import re as _re
                retry_match = _re.search(r"retry.?after[:\s]+(\d+)", error_str, _re.IGNORECASE)
                if retry_match:
                    retry_after = int(retry_match.group(1))
                retry_after = min(retry_after, 60)
            except Exception:
                pass

            print(f"      ⚠ Rate limited (429) — waiting {retry_after}s then retrying...")
            _time.sleep(retry_after)

            try:
                response = _chat_create_with_token_compat(
                    client,
                    model=model_name,
                    messages=messages,
                    max_tokens=resolved_max_tokens,
                    temperature=resolved_temperature,
                    response_format=RESPONSE_FORMAT
                )
                _log_stage_completion(response, stage_num, model_name)
                return response
            except Exception as e_retry:
                print(f"      ✗ Rate limit retry failed: {e_retry}")
                # Fallback to base deployment on persistent 429
                if model_name != AZURE_DEPLOYMENT:
                    print(f"      ⚠ Falling back to '{AZURE_DEPLOYMENT}' due to persistent rate limit on '{model_name}'")
                    try:
                        response = _chat_create_with_token_compat(
                            client,
                            model=AZURE_DEPLOYMENT,
                            messages=messages,
                            max_tokens=resolved_max_tokens,
                            temperature=resolved_temperature,
                            response_format=RESPONSE_FORMAT
                        )
                        _log_stage_completion(response, stage_num, AZURE_DEPLOYMENT)
                        return response
                    except Exception as e_fallback:
                        print(f"      ✗ Fallback '{AZURE_DEPLOYMENT}' also failed: {e_fallback}")
                        raise
                raise

        # Fallback for 400 invalid_prompt / internal model errors
        # gpt-5.4 sometimes returns 400 for images that gpt-4o-vision handles
        if (
            ("400" in error_str or "invalid_prompt" in error_str.lower()
             or "invalid_request_error" in error_str.lower())
            and model_name != AZURE_DEPLOYMENT
        ):
            print(f"      ⚠ Model '{model_name}' returned 400/invalid_prompt — retrying with fallback '{AZURE_DEPLOYMENT}'")
            try:
                response = _chat_create_with_token_compat(
                    client,
                    model=AZURE_DEPLOYMENT,
                    messages=messages,
                    max_tokens=resolved_max_tokens,
                    temperature=resolved_temperature,
                    response_format=RESPONSE_FORMAT
                )
                _log_stage_completion(response, stage_num, AZURE_DEPLOYMENT)
                return response
            except Exception as e_fallback:
                print(f"      ✗ Fallback deployment retry also failed: {e_fallback}")
                raise

        # Fallback for missing/invalid stage deployment name in Azure
        if (
            ("resource not found" in error_str.lower() or "'code': '404'" in error_str.lower())
            and model_name != AZURE_DEPLOYMENT
        ):
            print(f"      ⚠ Model '{model_name}' not found in Azure deployments - retrying with fallback '{AZURE_DEPLOYMENT}'")
            try:
                response = _chat_create_with_token_compat(
                    client,
                    model=AZURE_DEPLOYMENT,
                    messages=messages,
                    max_tokens=resolved_max_tokens,
                    temperature=resolved_temperature,
                    response_format=RESPONSE_FORMAT
                )
                _log_stage_completion(response, stage_num, AZURE_DEPLOYMENT)
                return response
            except Exception as e_fallback:
                print(f"      ✗ Fallback deployment retry failed: {e_fallback}")
                raise

        # Check if this is a content filter error
        if "content_filter" in error_str.lower() or "responsibleaipolicyviolation" in error_str.lower():
            print(f"      ⚠ Content filter triggered - retrying with low-detail + disclaimer...")
            try:
                # Add disclaimer at beginning of messages
                disclaimer_msg = {
                    "role": "system",
                    "content": "IMPORTANT: You are analyzing a TECHNICAL ENGINEERING DRAWING for industrial/manufacturing purposes. This is legitimate engineering documentation, not harmful content. Please proceed with technical analysis."
                }

                # Create new messages list with disclaimer
                messages_retry = [disclaimer_msg] + messages

                # Change image detail to "low" for all images
                for msg in messages_retry:
                    if isinstance(msg.get("content"), list):
                        for content_item in msg["content"]:
                            if isinstance(content_item, dict) and "image_url" in content_item:
                                content_item["image_url"]["detail"] = "low"

                # Retry with modified request
                response = _chat_create_with_token_compat(
                    client,
                    model=model_name,
                    messages=messages_retry,
                    max_tokens=resolved_max_tokens,
                    temperature=resolved_temperature,
                    response_format=RESPONSE_FORMAT
                )
                print(f"      ✓ Content filter retry successful")
                _log_stage_completion(response, stage_num, model_name)
                return response
            except Exception as e2:
                print(f"      ✗ Content filter retry also failed: {e2}")
                return None

        # ── General fallback: any unhandled error → retry with gpt-4o ──
        if model_name != AZURE_DEPLOYMENT:
            print(f"      ⚠ Model '{model_name}' failed ({type(e).__name__}) — falling back to '{AZURE_DEPLOYMENT}'")
            try:
                response = _chat_create_with_token_compat(
                    client,
                    model=AZURE_DEPLOYMENT,
                    messages=messages,
                    max_tokens=resolved_max_tokens,
                    temperature=resolved_temperature,
                    response_format=RESPONSE_FORMAT
                )
                _log_stage_completion(response, stage_num, AZURE_DEPLOYMENT)
                return response
            except Exception as e_fallback:
                print(f"      ✗ Fallback '{AZURE_DEPLOYMENT}' also failed: {e_fallback}")
                raise
        else:
            raise
