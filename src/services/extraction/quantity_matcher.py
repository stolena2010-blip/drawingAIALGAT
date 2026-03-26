"""
Quantity Matcher — Match quantities from orders/quotes/email to drawings
========================================================================

Extracted from customer_extractor_v3_dual.py (Step 5 of refactoring plan).

Functions:
  - match_quantities_to_drawings: Match quantities from orders/quotes/email
  - extract_base_and_suffix: Split P.N. into base + suffix
  - override_pn_from_email: Override P.N. with email-supplied suffix
"""

import re
from typing import Dict, List, Optional, Tuple

from src.services.extraction.filename_utils import _normalize_item_number
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _key_matches_any_drawing(eq_norm: str, drawing_norms: set) -> bool:
    """Return True if *eq_norm* plausibly refers to one of the drawings.

    Uses the same flexible matching as the main quantity loop:
    exact → containment → last-5-digit suffix.
    """
    if not eq_norm:
        return False
    for dn in drawing_norms:
        # exact or containment
        if eq_norm == dn or eq_norm in dn or dn in eq_norm:
            return True
        # suffix-digit match (≥5 digits each)
        eq_digits = re.sub(r'\D', '', eq_norm)
        dn_digits = re.sub(r'\D', '', dn)
        if len(eq_digits) >= 5 and len(dn_digits) >= 5:
            suf = min(5, len(eq_digits), len(dn_digits))
            if eq_digits[-suf:] == dn_digits[-suf:]:
                return True
    return False


def extract_base_and_suffix(pn: str) -> Tuple[str, Optional[str]]:
    """Split 'ABC-123-001' into ('ABC-123', '001'), or ('ABC-123', None).

    Also handles IAI format where suffix may end with a letter:
    'H2251-1941-003H' → ('H2251-1941', '003H')
    'H2251-0104-001H' → ('H2251-0104', '001H')
    """
    # Match 3-digit suffix optionally followed by a single letter (IAI: 001H, 003H)
    m = re.match(r'^(.+?)-(\d{3}[A-Za-z]?)$', pn.strip())
    if m:
        return m.group(1), m.group(2)
    return pn.strip(), None


def match_quantities_to_drawings(
    subfolder_results: List[Dict],
    item_details: Dict,
    email_data: Dict,
    general_quantities: List[str],
    pl_items_list: Optional[List[Dict]] = None,
) -> Tuple[int, int]:
    """
    Match quantities from orders/quotes/email to drawing results.

    Updates *subfolder_results* dicts in-place (adds ``quantity``,
    ``quantity_match_type``, ``quantity_source``, ``work_description_doc``,
    ``work_description_email``, ``email_from``, ``email_subject``,
    ``quantity_notes``).

    Returns
    -------
    (items_with_specific_qty, items_with_general_qty)
    """
    if pl_items_list is None:
        pl_items_list = []

    items_with_specific_qty = 0
    items_with_general_qty = 0

    # Ensure item_details is a dict before looping
    if not isinstance(item_details, dict):
        item_details = {}

    # Track which part_numbers already got quantities from orders/quotes
    parts_already_matched: Dict[str, list] = {}

    # ── Per-key validation: keep only part_quantities keys that match
    #    at least one drawing in the batch.  Keys that match nothing
    #    (phone numbers, spec words, etc.) are individually dropped.
    if email_data and email_data.get('part_quantities'):
        email_pq = email_data['part_quantities']
        drawing_norms = set()
        for rd in subfolder_results:
            pn = str(rd.get("part_number") or "").strip()
            if pn:
                drawing_norms.add(_normalize_item_number(pn))

        if drawing_norms:
            kept: dict = {}
            dropped: list = []
            for eq_key, eq_val in email_pq.items():
                eq_norm = _normalize_item_number(eq_key) if eq_key else ""
                if _key_matches_any_drawing(eq_norm, drawing_norms):
                    kept[eq_key] = eq_val
                else:
                    dropped.append(eq_key)
            if dropped:
                logger.info(
                    f"ℹ️ part_quantities: dropped {len(dropped)} unmatched key(s) "
                    f"{dropped} — kept {len(kept)} (drawings: {sorted(drawing_norms)})"
                )
                email_data['part_quantities'] = kept

    for idx_item, result_dict in enumerate(subfolder_results, 1):
        part_number = str(result_dict.get("part_number") or "").strip()
        item_name = str(result_dict.get("item_name") or "").strip()

        if not part_number:
            continue

        part_num_normalized = _normalize_item_number(part_number)

        quantity_str = ""
        work_desc_doc = ""
        work_desc_email = ""
        matched_key = None
        match_type = ""
        qty_source = ""

        # ── flexible matching — only with part_number ──
        # 1. exact match first
        if part_num_normalized in item_details:
            matched_key = part_num_normalized
        else:
            # 2. partial match (bidirectional containment)
            best_match = None
            best_match_score = 0
            part_digits = re.sub(r'\D', '', part_num_normalized)

            for key in item_details.keys():
                if part_num_normalized in key or key in part_num_normalized:
                    common_len = min(len(part_num_normalized), len(key))
                    if common_len >= 3:
                        if part_num_normalized in key:
                            score = len(part_num_normalized)
                        elif key in part_num_normalized:
                            score = len(key)
                        else:
                            score = 0

                        if score > best_match_score:
                            best_match = key
                            best_match_score = score
                else:
                    # fuzzy prefix check (≥70 %)
                    min_len = min(len(part_num_normalized), len(key))
                    if min_len >= 5:
                        compare_len = int(min_len * 0.7)
                        if part_num_normalized[:compare_len] == key[:compare_len]:
                            score = compare_len
                            if score > best_match_score:
                                best_match = key
                                best_match_score = score

                    # 3. digit-suffix comparison (OCR/AI errors)
                    if not best_match:
                        key_digits = re.sub(r'\D', '', key)
                        if len(part_digits) >= 5 and len(key_digits) >= 5:
                            suffix_len = min(5, len(part_digits), len(key_digits))
                            if part_digits[-suffix_len:] == key_digits[-suffix_len:]:
                                score = suffix_len
                                if score > best_match_score:
                                    best_match = key
                                    best_match_score = score

            if best_match:
                matched_key = best_match

        # ── apply match ──
        if matched_key:
            if part_num_normalized in parts_already_matched:
                prev_indices = parts_already_matched[part_num_normalized]

                if result_dict.get('validation_warnings'):
                    result_dict['validation_warnings'] += f" | כפילות: אותו מספר פריט כמו ציור #{prev_indices[0]}"
                else:
                    result_dict['validation_warnings'] = f"כפילות: אותו מספר פריט כמו ציור #{prev_indices[0]}"

                match_type = "כפילות - אותו פריט"
            else:
                match_type = "ספציפי למספר פריט"
                parts_already_matched[part_num_normalized] = []

            parts_already_matched[part_num_normalized].append(idx_item)

            details = item_details[matched_key]
            quantities = details.get('quantities', [])

            unique_quantities: list = []
            seen: set = set()
            for q in quantities:
                q_clean = str(q).strip().strip('()')
                if q_clean and q_clean not in seen:
                    unique_quantities.append(q_clean)
                    seen.add(q_clean)

            if len(unique_quantities) == 1:
                quantity_str = unique_quantities[0]
            elif len(unique_quantities) > 1:
                quantity_str = f"({', '.join(unique_quantities)})"

            work_desc_doc = details.get('work_description', '')
            qty_source = "הצעה/הזמנה"
            items_with_specific_qty += 1
            logger.info(f"Assigned from order/quote:")
            logger.info(f"Quantity: '{quantity_str}'")
            logger.info(f"Match Type: {match_type}")
            logger.info(f"Source: {qty_source}")
            logger.info(f"work_description_doc='{work_desc_doc}'")
        else:
            logger.info(f"NO MATCH: Could not match '{part_num_normalized}' with any item in orders/quotes")

            # Check if email has per-part quantities table
            # If yes, items NOT in the table should get qty=empty (not general)
            email_has_specific_qtys = bool(
                email_data and email_data.get('part_quantities')
                and len(email_data['part_quantities']) > 0
            )

            unmatched_quantities: list = []

            if not email_has_specific_qtys:
                # No per-part table in email → collect quantities from orders/quotes
                for item_key, details in item_details.items():
                    quantities = details.get('quantities', [])
                    for q in quantities:
                        q_clean = str(q).strip().strip('()')
                        if q_clean and q_clean not in unmatched_quantities:
                            unmatched_quantities.append(q_clean)

                if not unmatched_quantities and general_quantities:
                    for q in general_quantities:
                        q_clean = str(q).strip().strip('()')
                        if q_clean and q_clean not in unmatched_quantities:
                            unmatched_quantities.append(q_clean)
                    qty_source_unmatched = "מייל (כמויות כלליות)"
                else:
                    qty_source_unmatched = "הצעה/הזמנה (ללא התאמה לפריט)"

            if email_has_specific_qtys:
                # Email has a per-part quantities table but this item is NOT in it
                # → quantity should be empty (will try email matching below)
                quantity_str = ''
                match_type = ''
                qty_source = ''
                logger.info(
                    f"⚠️ Item '{part_num_normalized}' not in orders — "
                    f"email has per-part quantities, skipping general assignment"
                )
            elif unmatched_quantities:
                if len(unmatched_quantities) == 1:
                    quantity_str = unmatched_quantities[0]
                else:
                    quantity_str = f"({', '.join(unmatched_quantities)})"
                match_type = "לא התאמה - כמויות"
                qty_source = qty_source_unmatched
                logger.info(f"Assigning unmatched quantities:")
                logger.info(f"Quantities: {quantity_str}")
                logger.info(f"Match Type: {match_type}")
                logger.info(f"Source: {qty_source}")
            else:
                quantity_str = '0'
                match_type = "אין כמויות"
                qty_source = "לא מצא"

        # ── email part_quantities matching ──
        if (not quantity_str or not matched_key) and email_data and email_data.get('part_quantities'):
            email_parts = email_data.get('part_quantities', {})
            logger.info(f"Trying email part_quantities matching: drawing='{part_num_normalized}' vs email keys={list(email_parts.keys())[:10]}")

            if part_num_normalized in email_parts:
                qty_value = email_parts[part_num_normalized]
                if isinstance(qty_value, list):
                    cleaned_qtys = [str(q).strip().strip('()') for q in qty_value]
                    quantity_str = f"({', '.join(cleaned_qtys)})"
                else:
                    quantity_str = str(qty_value).strip().strip('()')

                match_type = "מייל (שורה)"
                qty_source = "מייל"
                items_with_general_qty += 1
                logger.info(f"Assigned quantity from email table:")
                logger.info(f"Quantity: '{quantity_str}'")
                logger.info(f"Match Type: {match_type}")
                logger.info(f"Source: {qty_source}")
            else:
                logger.info(f"No exact match in email, trying flexible matching...")
                best_email_match = None
                best_email_score = 0
                part_digits = re.sub(r'\D', '', part_num_normalized)

                for email_key in email_parts.keys():
                    email_key_normalized = _normalize_item_number(email_key)

                    if part_num_normalized in email_key_normalized or email_key_normalized in part_num_normalized:
                        common_len = min(len(part_num_normalized), len(email_key_normalized))
                        if common_len >= 3:
                            if part_num_normalized in email_key_normalized:
                                score = len(part_num_normalized)
                            elif email_key_normalized in part_num_normalized:
                                score = len(email_key_normalized)
                            else:
                                score = 0

                            if score > best_email_score:
                                best_email_match = email_key
                                best_email_score = score
                    else:
                        email_digits = re.sub(r'\D', '', email_key_normalized)
                        if len(part_digits) >= 5 and len(email_digits) >= 5:
                            suffix_len = min(5, len(part_digits), len(email_digits))
                            if part_digits[-suffix_len:] == email_digits[-suffix_len:]:
                                score = suffix_len + 1
                                if score > best_email_score:
                                    best_email_match = email_key
                                    best_email_score = score
                                    logger.info(f"Suffix match: '{part_num_normalized}' ~ '{email_key_normalized}' (last {suffix_len} digits)")

                if best_email_match:
                    qty_value = email_parts[best_email_match]
                    if isinstance(qty_value, list):
                        cleaned_qtys = [str(q).strip().strip('()') for q in qty_value]
                        quantity_str = f"({', '.join(cleaned_qtys)})"
                    else:
                        quantity_str = str(qty_value).strip().strip('()')

                    match_type = "מייל (שורה - התאמה גמישה)"
                    qty_source = "מייל"
                    items_with_general_qty += 1
                    logger.info(f"Assigned quantity from email (flexible match):")
                    logger.info(f"Matched '{part_num_normalized}' with '{best_email_match}'")
                    logger.info(f"Quantity: '{quantity_str}'")
                    logger.info(f"Match Type: {match_type}")
                    logger.info(f"Source: {qty_source}")

        # ── general email quantities fallback ──
        # Assign general quantity as fallback for items without a specific match.
        # When email has BOTH per-part quantities AND a general quantity
        # (e.g. "2143: 120 יח, כל השאר: 30 יח"), the general quantity serves
        # as default for items not listed specifically.
        email_has_specific_qtys_fallback = bool(
            email_data and email_data.get('part_quantities')
            and len(email_data['part_quantities']) > 0
        )
        if not quantity_str and general_quantities:
            unique_general: list = []
            seen_general: set = set()
            for q in general_quantities:
                q_clean = str(q).strip().strip('()')
                if q_clean and q_clean not in seen_general:
                    unique_general.append(q_clean)
                    seen_general.add(q_clean)

            if len(unique_general) == 1:
                quantity_str = unique_general[0]
            elif len(unique_general) > 1:
                quantity_str = f"({', '.join(unique_general)})"
            match_type = "כללי"
            qty_source = "מייל"
            items_with_general_qty += 1
            logger.info(f"Assigned general quantity from email:")
            logger.info(f"Quantity: '{quantity_str}'")
            logger.info(f"Match Type: {match_type}")
            logger.info(f"Source: {qty_source}")
        elif not quantity_str and email_has_specific_qtys_fallback and not general_quantities:
            # Email has per-part table but NO general fallback → qty empty
            quantity_str = ''
            match_type = 'לא נמצא ברשימת הכמויות'
            qty_source = 'לא נמצא במייל'
            logger.info(
                f"⚠️ Item '{part_num_normalized}' not found in email quantities table "
                f"— qty set to empty (other items have specific quantities)"
            )

        # ── work_description_email: from AI-extracted email fields ──
        if email_data:
            # 1. Per-part work description from AI (highest priority)
            part_work_descs = email_data.get('part_work_descriptions', {})
            matched_wd_key = None
            if part_work_descs:
                # Exact match first
                if part_num_normalized in part_work_descs:
                    matched_wd_key = part_num_normalized
                else:
                    # Flexible matching (same logic as quantity matching)
                    best_wd_match = None
                    best_wd_score = 0
                    part_digits_wd = re.sub(r'\D', '', part_num_normalized)

                    for wd_key in part_work_descs.keys():
                        # Bidirectional containment
                        if part_num_normalized in wd_key or wd_key in part_num_normalized:
                            common_len = min(len(part_num_normalized), len(wd_key))
                            if common_len >= 3:
                                score = len(part_num_normalized) if part_num_normalized in wd_key else len(wd_key)
                                if score > best_wd_score:
                                    best_wd_match = wd_key
                                    best_wd_score = score
                        else:
                            # Digit-suffix comparison (OCR/AI errors)
                            wd_key_digits = re.sub(r'\D', '', wd_key)
                            if len(part_digits_wd) >= 5 and len(wd_key_digits) >= 5:
                                suffix_len = min(5, len(part_digits_wd), len(wd_key_digits))
                                if part_digits_wd[-suffix_len:] == wd_key_digits[-suffix_len:]:
                                    score = suffix_len
                                    if score > best_wd_score:
                                        best_wd_match = wd_key
                                        best_wd_score = score

                    if best_wd_match:
                        matched_wd_key = best_wd_match
                        logger.info(f"work_description_email flexible match: '{part_num_normalized}' ~ '{best_wd_match}'")

            if matched_wd_key:
                work_desc_email = part_work_descs[matched_wd_key]
                logger.info(f"work_description_email (per-part AI): '{work_desc_email[:60]}'")
            else:
                # If email has per-part descriptions, don't assign general to unmatched items
                email_has_specific_descs = bool(part_work_descs and len(part_work_descs) > 0)

                if email_has_specific_descs:
                    # Per-part descriptions exist but this item wasn't matched
                    # → leave empty (don't assign generic description)
                    logger.info(
                        f"⚠️ work_description_email: item '{part_num_normalized}' not matched "
                        f"in per-part descriptions ({len(part_work_descs)} items) — leaving empty"
                    )
                else:
                    # 2. General work description from AI
                    general_wd = email_data.get('general_work_description', '')
                    if general_wd:
                        work_desc_email = general_wd
                        # Append negation if exists
                        negation = email_data.get('work_description_negation', '')
                        if negation:
                            work_desc_email = f"{work_desc_email} | {negation}"
                        logger.info(f"work_description_email (general AI): '{work_desc_email[:60]}'")
                    else:
                        # 3. Fallback: keyword-based extraction from email body (legacy)
                        wd_keyword = email_data.get('work_description', '')
                        if wd_keyword:
                            work_desc_email = wd_keyword
                            logger.info(f"work_description_email (keyword fallback): '{work_desc_email[:60]}'")
                        # 4. Negation-only
                        elif email_data.get('work_description_negation', ''):
                            work_desc_email = email_data['work_description_negation']
                            logger.info(f"work_description_email (negation only): '{work_desc_email[:60]}'")

        # ── update result ──
        result_dict['quantity'] = quantity_str
        result_dict['quantity_match_type'] = match_type
        result_dict['quantity_source'] = qty_source
        result_dict['work_description_doc'] = work_desc_doc
        result_dict['work_description_email'] = work_desc_email
        result_dict['email_from'] = email_data.get('from', '') if email_data else ''
        result_dict['email_subject'] = email_data.get('subject', '') if email_data else ''

        if email_data and email_data.get('quantity_summary'):
            result_dict['quantity_notes'] = email_data.get('quantity_summary', '')

        # Update PL items with matched_drawing (if matched)
        if matched_key:
            drawing_name = result_dict.get('file_name', '')
            for pl_item in pl_items_list:
                if (pl_item['part_number'].startswith(part_num_normalized) or
                        part_num_normalized.startswith(pl_item['part_number'])):
                    pl_item['matched_drawing'] = drawing_name

    return items_with_specific_qty, items_with_general_qty


# ── IAI prefix normalization for base comparison ──────────────────────
_IAI_PREFIX_RE = re.compile(r'^(?:MD-H|MD-|H(?=[0-9]))', re.IGNORECASE)


def _normalize_iai_base(base: str) -> str:
    """Strip IAI-specific prefixes for base comparison.

    'MD-H2251-0104' → '2251-0104'
    'H2251-1941'    → '2251-1941'
    '2251-0104'     → '2251-0104'  (no-op)
    """
    return _IAI_PREFIX_RE.sub('', base)


def override_pn_from_email(
    subfolder_results: List[Dict],
    email_data: Dict,
    is_iai: bool = False,
) -> int:
    """
    Override P.N. with email-supplied suffix when a single unambiguous
    match is found.

    For IAI drawings the email often carries the authoritative part number
    (with correct suffix like ``-001H``).  When *is_iai* is True an extra
    matching pass strips IAI prefixes (``MD-H``, ``H``) so that
    ``MD-H2251-0104`` (drawing) matches ``2251-0104-001H`` (email).

    Returns the number of overrides applied.
    """
    if not email_data or not email_data.get('part_quantities'):
        return 0

    email_parts_raw = list(email_data['part_quantities'].keys())
    override_count = 0

    for result_dict in subfolder_results:
        current_pn = str(result_dict.get('part_number', '')).strip()
        if not current_pn:
            continue

        current_base, current_suffix = extract_base_and_suffix(current_pn)

        # ── Pass 1: direct base comparison (works for all customers) ──
        email_matches: list = []
        for ek_raw in email_parts_raw:
            ek_base, ek_suffix = extract_base_and_suffix(ek_raw)
            if ek_base.lower() == current_base.lower() and ek_suffix is not None:
                email_matches.append(ek_raw.strip())

        # ── Pass 2 (IAI only): compare after stripping IAI prefixes ──
        if not email_matches and is_iai:
            current_base_norm = _normalize_iai_base(current_base).lower()
            # Also try matching with the full PN as base (no suffix split)
            current_pn_norm = _normalize_iai_base(current_pn).lower()
            for ek_raw in email_parts_raw:
                ek_base, ek_suffix = extract_base_and_suffix(ek_raw)
                ek_base_norm = _normalize_iai_base(ek_base).lower()
                if ek_suffix is not None and (
                    ek_base_norm == current_base_norm
                    or ek_base_norm == current_pn_norm
                ):
                    email_matches.append(ek_raw.strip())

        if len(email_matches) == 1:
            new_pn = email_matches[0]
            if new_pn != current_pn:
                old_pn = current_pn
                if not result_dict.get('part_number_ocr_original'):
                    result_dict['part_number_ocr_original'] = old_pn
                result_dict['part_number'] = new_pn
                result_dict['pl_part_number'] = new_pn
                result_dict['pl_override_note'] = f'ממייל (מקור: {old_pn})'
                logger.info(f"✓ EMAIL OVERRIDE: '{old_pn}' → '{new_pn}' (from email body)")
                override_count += 1
        elif len(email_matches) > 1:
            logger.info(f"⚠️ EMAIL: multiple suffix matches for '{current_pn}': {email_matches} — not overriding")

    return override_count
