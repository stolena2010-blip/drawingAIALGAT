"""
Stage 9 — Smart Description Merge (o4-mini)
============================================
Merges 5 description sources per drawing item into 4 structured fields:
  merged_processes, merged_specs, merged_bom, merged_notes

Uses o4-mini (text-only reasoning model) — no Vision needed.
"""

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openai import AzureOpenAI

from src.services.ai.vision_api import (
    _chat_create_with_token_compat,
    _resolve_stage_call_config,
    _calculate_stage_cost,
)
logger = logging.getLogger(__name__)

STAGE_NUM = 9
BATCH_SIZE = 10  # items per API call (conservative for token budget)

# ── prompt loading ──────────────────────────────────────────────
_PROMPT_DIR = Path(__file__).resolve().parents[3] / "prompts"


def _load_system_prompt() -> str:
    prompt_path = _PROMPT_DIR / "09_merge_work_descriptions.txt"
    if prompt_path.exists():
        return prompt_path.read_text(encoding="utf-8").strip()
    logger.warning(f"[STAGE9] Prompt file not found: {prompt_path}")
    return ""


# ── helpers ─────────────────────────────────────────────────────

_DESC_FIELDS = (
    "process_summary_hebrew",
    "specifications",
    "PL Summary Hebrew",
    "work_description_email",
    "work_description_doc",
)


def _has_any_description(item: Dict) -> bool:
    """Return True if the item has at least one non-empty description source."""
    return any(
        str(item.get(f) or "").strip()
        for f in _DESC_FIELDS
    )


def _build_items_text(items: List[Dict]) -> str:
    """Build the user-message text listing all items for the AI."""
    parts: list[str] = []
    for idx, item in enumerate(items):
        pn = item.get("part_number", f"item_{idx}")
        parts.append(f"\n--- פריט #{idx}: {pn} ---")
        parts.append(f"שרטוט: {item.get('process_summary_hebrew') or 'ריק'}")
        parts.append(f"מפרטים (שרטוט): {item.get('specifications') or 'ריק'}")
        parts.append(f"PL: {item.get('PL Summary Hebrew') or 'ריק'}")
        parts.append(f"מייל: {item.get('work_description_email') or 'ריק'}")
        parts.append(f"מסמך: {item.get('work_description_doc') or 'ריק'}")
    return "\n".join(parts)


# ── Structured BOM builder (replaces AI bom field) ────────────

def _format_insert_entry(cat: str, qty: str, price, currency: str) -> str:
    """Format a single insert: PN ×qty ×price currency."""
    entry = f"{cat} ×{qty}" if cat and qty else cat
    if price is not None:
        entry += f" ×{price}{currency}"
    return entry


def _count_pl_primary_types(pl_hw: str) -> int:
    """Count primary insert types in PL Hardware string.

    Groups are separated by ``|``.  Within each group, alternates are
    marked with ``(חלופי)`` after a comma.  Each group = 1 primary type.
    """
    if not pl_hw:
        return 0
    return sum(1 for g in pl_hw.split('|') if g.strip())


def _build_structured_bom(item: Dict) -> str:
    """
    Build merged_bom from structured drawing inserts + PL Hardware.

    Output format::

        קשיחים [2]:
        שרטוט: K500-001 ×12 ×0.45₪ | MS51835 ×4
        PL: K500-001 ×12 ×0.45₪, MS124655 (חלופי) ×12 | MS51831-203 ×4 ×7.0₪

    If only one source has data, returns a single line:
        קשיחים [2]: K500-001 ×12 ×0.45₪ | MS51835 ×4

    The count ``[N]`` shows how many *primary* (non-alternate) insert types
    are required.  If both sources exist and disagree, the PL count wins.
    """
    lines: list[str] = []

    # — Drawing inserts (from inserts_hardware list-of-dicts) —
    hw_list = item.get('inserts_hardware')
    drawing_parts: list[str] = []
    if isinstance(hw_list, list):
        for hw in hw_list:
            if not isinstance(hw, dict):
                continue
            cat = str(hw.get('cat_no') or '').strip()
            qty = str(hw.get('qty') or '').strip()
            if not cat:
                continue
            price = hw.get('unit_price')
            currency = str(hw.get('currency') or '')
            drawing_parts.append(_format_insert_entry(cat, qty, price, currency))
    elif isinstance(hw_list, str) and hw_list.strip():
        # Already formatted (shouldn't happen at stage 9 time, but be safe)
        drawing_parts.append(hw_list.strip())

    # — PL inserts (pre-formatted string with prices & חלופי markers) —
    pl_hw = str(item.get('PL Hardware') or '').strip()

    # — Determine primary-type count —
    has_drawing = bool(drawing_parts)
    has_pl = bool(pl_hw)
    drawing_count = len(drawing_parts)          # all drawing inserts are primary
    pl_count = _count_pl_primary_types(pl_hw)   # pipe-separated groups

    if has_drawing and has_pl:
        # Discrepancy → PL wins
        type_count = pl_count if pl_count != drawing_count else drawing_count
    elif has_pl:
        type_count = pl_count
    else:
        type_count = drawing_count

    # — Build output —
    if has_drawing and has_pl:
        lines.append(f"קשיחים [{type_count}]:")
        lines.append("שרטוט: " + " | ".join(drawing_parts))
        lines.append("עץ: " + pl_hw)
    elif has_drawing:
        lines.append(f"קשיחים [{type_count}]: " + " | ".join(drawing_parts))
    elif has_pl:
        lines.append(f"קשיחים [{type_count}]: " + pl_hw)

    return "\n".join(lines)


def _sum_pl_primary_qty(pl_hw: str) -> int:
    """Sum total quantity of *primary* inserts from a PL Hardware string.

    Groups are separated by ``|``.  Within each group the first entry is
    primary; entries containing ``(חלופי)`` are alternates and skipped.
    Quantity is extracted from the ``×<qty>`` token.

    Example::

        "K500 ×4 ×0.45₪, ALT1 (חלופי) ×4 | MS51835 ×8"
        → primary entries: K500 ×4, MS51835 ×8 → total = 12
    """
    if not pl_hw:
        return 0
    total = 0
    for group in pl_hw.split('|'):
        group = group.strip()
        if not group:
            continue
        # First comma-separated part is the primary; rest are alternates
        primary_part = group.split(',')[0].strip()
        if '(חלופי)' in primary_part:
            continue  # safety: skip if somehow the first part is an alternate
        # Extract quantity from ×<number> tokens (first ×N is qty)
        qty_matches = re.findall(r'×(\d+(?:\.\d+)?)', primary_part)
        if qty_matches:
            try:
                total += int(float(qty_matches[0]))
            except (ValueError, IndexError):
                pass
    return total


def _sum_drawing_primary_qty(hw_list) -> int:
    """Sum total quantity of inserts from the drawing inserts_hardware list."""
    if not isinstance(hw_list, list):
        return 0
    total = 0
    for hw in hw_list:
        if not isinstance(hw, dict):
            continue
        cat = str(hw.get('cat_no') or '').strip()
        if not cat:
            continue
        qty_str = str(hw.get('qty') or '').strip()
        if qty_str:
            try:
                total += int(float(qty_str))
            except (ValueError, TypeError):
                pass
    return total


def _calc_hardware_count(item: Dict) -> int:
    """Calculate total primary hardware quantity.

    Priority: PL wins over drawing when both exist.
    Only counts primary inserts (skips חלופי alternates).
    """
    pl_hw = str(item.get('PL Hardware') or '').strip()
    hw_list = item.get('inserts_hardware')
    has_pl = bool(pl_hw)
    has_drawing = isinstance(hw_list, list) and any(
        isinstance(h, dict) and str(h.get('cat_no') or '').strip()
        for h in hw_list
    )

    if has_pl:
        return _sum_pl_primary_qty(pl_hw)
    elif has_drawing:
        return _sum_drawing_primary_qty(hw_list)
    return 0


def _build_merged_description(item: Dict) -> str:
    """Combine merged_processes + H.C (hardware count) into one field.

    Format::

        (תהליכים) אלומיניום 5052 | אנודייז | צביעה | H.C=19

    - Processes are kept as-is.
    - Notes (merged_notes) are excluded.
    - BOM details (merged_bom) are excluded; replaced by total hardware
      count appended after the last process with a ``|`` separator.
    - ``H.C=N`` where N = sum of primary insert quantities
      (PL source wins over drawing when both exist).
    - If there are no inserts, H.C is omitted entirely.
    """
    procs = str(item.get('merged_processes') or '').strip()
    if not procs or procs == 'nan':
        # No processes — build from available fields
        parts: list[str] = []
        specs = str(item.get('merged_specs') or '').strip()
        if specs and specs != 'nan':
            parts.append(f"(מפרטים) {specs}")
        notes = str(item.get('merged_notes') or '').strip()
        if notes and notes != 'nan':
            parts.append(f"(הערות) {notes}")
        hc = _calc_hardware_count(item)
        if hc:
            parts.append(f"H.C={hc}")
        if not parts:
            # Last resort: use item_name if available
            item_name = str(item.get('item_name') or '').strip()
            if item_name and item_name != 'nan':
                return item_name
        return ' | '.join(parts)

    hc = _calc_hardware_count(item)
    if hc:
        return f"(תהליכים) {procs} | H.C={hc}"
    return f"(תהליכים) {procs}"


# ── core API call ───────────────────────────────────────────────

def _call_merge_api(
    items: List[Dict],
    client: AzureOpenAI,
) -> Tuple[List[Dict], int, int]:
    """
    Send a batch of items to o4-mini and return parsed results.

    Returns (results_list, tokens_in, tokens_out).
    Each result dict has keys: idx, processes, specs, bom, notes.
    """
    system_prompt = _load_system_prompt()
    if not system_prompt:
        return [], 0, 0

    items_text = _build_items_text(items)
    user_message = f"הפריטים:\n{items_text}"

    model, max_tokens, temperature = _resolve_stage_call_config(
        STAGE_NUM, default_max_tokens=32000, default_temperature=0,
    )

    try:
        response = _chat_create_with_token_compat(
            client,
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message},
            ],
            max_tokens=max_tokens,
            temperature=temperature,
        )
    except Exception as e:
        logger.error(f"[STAGE9] API call failed: {e}")
        return [], 0, 0

    tokens_in = getattr(response.usage, "prompt_tokens", 0) or 0
    tokens_out = getattr(response.usage, "completion_tokens", 0) or 0

    raw = (response.choices[0].message.content or "").strip()
    if not raw:
        logger.warning(
            f"[STAGE9] Empty response from AI "
            f"(tokens: {tokens_in:,} in / {tokens_out:,} out — "
            f"reasoning model may have exhausted budget on internal thinking)"
        )
        return [], tokens_in, tokens_out

    # Parse JSON — strip markdown fences if present
    json_text = raw
    if json_text.startswith("```"):
        first_nl = json_text.find("\n")
        if first_nl != -1:
            json_text = json_text[first_nl + 1:]
        if json_text.rstrip().endswith("```"):
            json_text = json_text.rstrip()[:-3].rstrip()

    # Try parsing as JSON array
    try:
        results = json.loads(json_text)
        if isinstance(results, dict) and "items" in results:
            results = results["items"]
        if not isinstance(results, list):
            results = [results]
        return results, tokens_in, tokens_out
    except json.JSONDecodeError:
        pass

    # Retry: try to find JSON array within the text
    match = re.search(r"\[[\s\S]*\]", json_text)
    if match:
        try:
            results = json.loads(match.group())
            return results, tokens_in, tokens_out
        except json.JSONDecodeError:
            pass

    logger.warning(f"[STAGE9] Failed to parse JSON response ({len(raw)} chars)")
    logger.debug(f"[STAGE9] Response: {raw[:300]}")
    return [], tokens_in, tokens_out


# ── public entry point ──────────────────────────────────────────

def merge_descriptions(
    subfolder_results: List[Dict],
    client: AzureOpenAI,
) -> Tuple[int, int, int]:
    """
    Stage 9: Merge descriptions for all items in subfolder_results.

    Updates each result dict in-place with:
        merged_processes, merged_specs, merged_bom, merged_notes

    Items with no description sources are skipped (fields set to '').

    Returns (items_merged, total_tokens_in, total_tokens_out).
    """
    # Initialize fields for ALL items
    for r in subfolder_results:
        r["merged_processes"] = ""
        r["merged_specs"] = ""
        r["merged_bom"] = ""
        r["merged_notes"] = ""

    # Filter to items that have at least one description
    items_with_desc: List[Tuple[int, Dict]] = [
        (i, r) for i, r in enumerate(subfolder_results) if _has_any_description(r)
    ]

    if not items_with_desc:
        logger.info("[STAGE9] No items with descriptions — skipping merge")
        return 0, 0, 0

    total_tokens_in = 0
    total_tokens_out = 0
    total_merged = 0

    # ── Helper: process a single batch (with retry-split on empty) ──
    def _process_batch(batch, batch_label: str) -> None:
        nonlocal total_tokens_in, total_tokens_out, total_merged

        batch_items = [item for _, item in batch]
        results, tok_in, tok_out = _call_merge_api(batch_items, client)
        total_tokens_in += tok_in
        total_tokens_out += tok_out

        cost = _calculate_stage_cost(tok_in, tok_out, STAGE_NUM)
        logger.info(
            f"[STAGE9] {batch_label}: "
            f"{len(batch_items)} items → {len(results)} results "
            f"(${cost:.4f})"
        )

        # If the model returned 0 results and batch has >1 item, split and retry
        if len(results) == 0 and len(batch) > 1:
            mid = len(batch) // 2
            logger.warning(
                f"[STAGE9] {batch_label} returned 0 results — "
                f"splitting into {mid} + {len(batch) - mid} items and retrying"
            )
            _process_batch(batch[:mid], f"{batch_label}a")
            _process_batch(batch[mid:], f"{batch_label}b")
            return

        # Single-item batch returned 0 results → retry once
        if len(results) == 0 and len(batch) == 1 and not batch_label.endswith('_retry'):
            logger.warning(
                f"[STAGE9] {batch_label} single-item returned 0 results — retrying once"
            )
            _process_batch(batch, f"{batch_label}_retry")
            return

        # Map results back by index
        for result_idx, merged in enumerate(results):
            if result_idx >= len(batch):
                break  # AI returned more items than sent

            original_idx, original_item = batch[result_idx]

            processes = merged.get("processes") or ""
            specs = merged.get("specs") or ""
            notes = merged.get("notes") or ""

            subfolder_results[original_idx]["merged_processes"] = str(processes).strip()
            subfolder_results[original_idx]["merged_specs"] = str(specs).strip()
            subfolder_results[original_idx]["merged_notes"] = str(notes).strip()
            total_merged += 1

    # Process in batches
    for batch_start in range(0, len(items_with_desc), BATCH_SIZE):
        batch = items_with_desc[batch_start : batch_start + BATCH_SIZE]
        batch_num = batch_start // BATCH_SIZE + 1
        _process_batch(batch, f"Batch {batch_num}")

    # ── Build merged_bom programmatically (structured, not AI) ──
    for i, r in enumerate(subfolder_results):
        r["merged_bom"] = _build_structured_bom(r)

    # ── Build merged_description (combined processes + bom + notes) ──
    for r in subfolder_results:
        r["merged_description"] = _build_merged_description(r)

    # ── Color / paint price lookup ──
    from src.services.extraction.color_price_lookup import lookup_color_prices
    for r in subfolder_results:
        r["color_prices"] = lookup_color_prices(
            r.get("merged_specs", ""),
            r.get("merged_processes", ""),
        )

    logger.info(
        f"[STAGE9] Merged {total_merged}/{len(items_with_desc)} items "
        f"(tokens: {total_tokens_in:,} in / {total_tokens_out:,} out)"
    )
    return total_merged, total_tokens_in, total_tokens_out
