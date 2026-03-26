"""
P.N. / Drawing# sanity checks, disambiguation, and confidence calculation.
============================================================================

Extracted from customer_extractor_v3_dual.py for modularity.

Functions:
- is_cage_code(): Check if value looks like a CAGE code
- run_pn_sanity_checks(): All P.N./Drawing# sanity checks A-D and fallbacks
- calculate_confidence(): Confidence level based on filename matching
"""

import re
from pathlib import Path
from typing import Optional

from src.services.extraction.filename_utils import (
    check_value_in_filename,
    check_exact_match_in_filename,
    _disambiguate_part_number,
    extract_part_number_from_filename,
)
from src.services.extraction.pn_voting import extract_pn_dn_from_text
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _find_near_match_in_filename(ocr_value: str, filename: str) -> Optional[str]:
    """
    Find a near-match (exactly 1 char different) in the filename.
    Catches ALL single-character misreads: 5↔6, O↔0, 8↔B, etc.

    Builds both individual tokens and compound segments (e.g. '0317-055')
    from the filename, normalizes both sides, and compares.

    Returns the filename version if found, otherwise None.
    """
    if not ocr_value or not filename:
        return None

    ocr_norm = re.sub(r'[^A-Za-z0-9]', '', str(ocr_value)).upper()
    if len(ocr_norm) < 4:
        return None

    fn_stem = Path(filename).stem if '.' in filename else filename

    # Collect candidate segments from filename
    candidates = []  # list of (display_form, normalized_form)

    # 1. Compound segments: tokens joined by single separators (e.g. '0317-055')
    for m in re.finditer(r'[A-Za-z0-9]+(?:[-_.][A-Za-z0-9]+)+', fn_stem):
        seg = m.group()
        candidates.append((seg, re.sub(r'[^A-Za-z0-9]', '', seg).upper()))

    # 2. Individual alphanumeric tokens
    for tok in re.findall(r'[A-Za-z0-9]+', fn_stem):
        tok_norm = tok.upper()
        if len(tok_norm) >= 4:
            candidates.append((tok, tok_norm))

    # Search for 1-char-diff match
    for display, cand_norm in candidates:
        if len(cand_norm) != len(ocr_norm):
            continue
        diff_count = sum(1 for a, b in zip(ocr_norm, cand_norm) if a != b)
        if diff_count == 1:
            return display

    return None


def is_cage_code(value) -> bool:
    """Check if value matches CAGE CODE format (5-6 alphanumeric chars, mixed letters+digits)"""
    if not value or not isinstance(value, str):
        return False
    value = str(value).strip()
    # CAGE CODE can be 5 OR 6 characters
    if len(value) in [5, 6] and value.isalnum():
        has_letter = bool(re.search(r'[A-Z]', value.upper()))
        has_digit = bool(re.search(r'[0-9]', value))
        return has_letter and has_digit
    return False


def run_pn_sanity_checks(
    result_data: dict,
    filename: str,
    file_path: str,
    pdfplumber_text: str = '',
    is_rafael: bool = False,
    is_iai: bool = False,
) -> dict:
    """
    Run all P.N./Drawing# sanity checks A-E and fallbacks.
    Modifies result_data in place and returns it.

    Includes:
    - Smart disambiguation for P.N. and Drawing#
    - P.N. extend from filename
    - CAGE code removal
    - Sanity checks A-D (P.N. checks)
    - Sanity check E (D.N.: Drawing# not in filename but P.N. is → DN=PN)
    - Fallback copy between P.N. and Drawing#
    - FREE fallbacks (PL Part Number, filename extraction)
    - IAI unification rule
    """
    # ============ REJECT DATE VALUES AS PART/DRAWING NUMBERS ============
    # Vision sometimes extracts DATE fields (01.02.24, 04.02.2024) as identifiers
    _DATE_PATTERN = re.compile(
        r'^(?:'
        r'\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}'
        r'|\d{4}[./\-]\d{1,2}[./\-]\d{1,2}'
        r')$'
    )
    for _field in ('part_number', 'drawing_number'):
        val = str(result_data.get(_field, '') or '').strip()
        if val and _DATE_PATTERN.match(val):
            logger.warning(f"⚠️  {_field} '{val}' is a DATE — clearing")
            result_data[_field] = None

    # ============ SMART DISAMBIGUATION FOR PART NUMBER ============
    if result_data.get('part_number'):
        ocr_part = result_data['part_number']

        # QUICK CHECK: Does it already match filename exactly? Skip if yes (optimize)
        normalized_ocr = re.sub(r'[^A-Za-z0-9]', '', str(ocr_part)).upper()
        filename_tokens = re.findall(r'[A-Za-z0-9]+', filename)

        has_exact_match = any(
            normalized_ocr == re.sub(r'[^A-Za-z0-9]', '', token).upper()
            for token in filename_tokens
        )

        if has_exact_match:
            logger.info(f"Part# perfect match: '{ocr_part}' ✓")
        else:
            disambiguation_result = _disambiguate_part_number(ocr_part, filename)
            best_candidate = disambiguation_result['best_candidate']
            confidence = disambiguation_result['confidence']
            method = disambiguation_result['method']

            if best_candidate != ocr_part:
                logger.info(f"Part# corrected ({method}): '{ocr_part}' → '{best_candidate}' ({confidence:.1%})")
                result_data['part_number'] = best_candidate
            else:
                # Last resort: filename has a near-match (1 char different)?
                near_match = _find_near_match_in_filename(ocr_part, filename)
                if near_match:
                    logger.info(f"Part# near-match fix: '{ocr_part}' → '{near_match}' (filename preferred, 1 char diff)")
                    result_data['part_number'] = near_match
                else:
                    logger.info(f"Part# no match in filename, using OCR: '{ocr_part}'")

    # ============ SMART DISAMBIGUATION FOR DRAWING NUMBER ============
    if result_data.get('drawing_number'):
        ocr_dwg = result_data['drawing_number']

        normalized_ocr = re.sub(r'[^A-Za-z0-9]', '', str(ocr_dwg)).upper()
        filename_tokens = re.findall(r'[A-Za-z0-9]+', filename)

        has_exact_match = any(
            normalized_ocr == re.sub(r'[^A-Za-z0-9]', '', token).upper()
            for token in filename_tokens
        )

        if has_exact_match:
            logger.info(f"Drawing# perfect match: '{ocr_dwg}' ✓")
        else:
            disambiguation_result = _disambiguate_part_number(ocr_dwg, filename)
            best_candidate = disambiguation_result['best_candidate']
            confidence = disambiguation_result['confidence']
            method = disambiguation_result['method']

            if best_candidate != ocr_dwg:
                logger.info(f"Drawing# corrected ({method}): '{ocr_dwg}' → '{best_candidate}' ({confidence:.1%})")
                result_data['drawing_number'] = best_candidate
            else:
                # Last resort: filename has a near-match (1 char different)?
                near_match = _find_near_match_in_filename(ocr_dwg, filename)
                if near_match:
                    logger.info(f"Drawing# near-match fix: '{ocr_dwg}' → '{near_match}' (filename preferred, 1 char diff)")
                    result_data['drawing_number'] = near_match
                else:
                    logger.info(f"Drawing# no match in filename, using OCR: '{ocr_dwg}'")

    # ─── P.N. EXTEND FROM FILENAME ───
    # If part_number is a prefix of the filename but not the full match,
    # extend it to the full version from the filename.
    # Example: P.N.='ET02-PF-16' filename='ET02-PF-16-03' → P.N.='ET02-PF-16-03'
    if result_data.get('part_number'):
        pn = str(result_data['part_number']).strip()
        fn_stem = Path(file_path).stem
        fn_clean = fn_stem.upper().replace('_', '-')
        pn_clean = pn.upper().replace('_', '-')

        if len(pn_clean) >= 6 and pn_clean != fn_clean:
            if fn_clean.startswith(pn_clean) and len(fn_clean) > len(pn_clean):
                pattern = re.escape(pn).replace(r'\-', '[-_]')
                match = re.search(pattern + r'[-_](\d{1,4})', fn_stem, re.IGNORECASE)
                if match:
                    extended_pn = fn_stem[match.start():match.end()]
                    logger.info(
                        f"📐 P.N. EXTEND: '{pn}' → '{extended_pn}' "
                        f"(filename has full version)"
                    )
                    result_data['part_number'] = extended_pn

    # ─── FINAL VALIDATION: CAGE CODE removal ───
    if result_data.get('part_number'):
        part_in_filename = check_value_in_filename(result_data['part_number'], filename)
        if not part_in_filename and is_cage_code(result_data['part_number']):
            logger.info(f"FINAL CHECK: Part# '{result_data['part_number']}' is CAGE CODE (not in filename) - removed!")
            result_data['part_number'] = None

    if result_data.get('drawing_number'):
        drawing_in_filename = check_value_in_filename(result_data['drawing_number'], filename)
        if not drawing_in_filename and is_cage_code(result_data['drawing_number']):
            logger.info(f"FINAL CHECK: Drawing# '{result_data['drawing_number']}' is CAGE CODE (not in filename) - removed!")
            result_data['drawing_number'] = None

    # ─── P.N. CLEANUP: Remove REV suffix if accidentally included ───
    # GPT sometimes includes revision in part_number: "FT-15912029-00-REVA"
    # Pattern: PN ends with -REV followed by letter(s) or digit(s)
    if result_data.get('part_number'):
        pn = str(result_data['part_number']).strip()
        rev = str(result_data.get('revision', '')).strip()

        rev_suffix_match = re.search(r'[-_](REV\.?\s*[A-Z0-9]+)$', pn, re.IGNORECASE)
        if rev_suffix_match:
            extracted_rev = rev_suffix_match.group(1)
            cleaned_pn = pn[:rev_suffix_match.start()]
            logger.warning(
                f"⚠️  P.N. CLEANUP: Removed REV suffix from P.N. "
                f"'{pn}' → '{cleaned_pn}' (extracted REV: '{extracted_rev}')"
            )
            result_data['part_number'] = cleaned_pn
            # If revision was empty, use the extracted one
            if not rev:
                rev_clean = re.sub(r'^REV\.?\s*', '', extracted_rev, flags=re.IGNORECASE)
                if rev_clean:
                    result_data['revision'] = rev_clean

    # Same cleanup for drawing_number
    if result_data.get('drawing_number'):
        dn = str(result_data['drawing_number']).strip()
        rev_suffix_match = re.search(r'[-_](REV\.?\s*[A-Z0-9]+)$', dn, re.IGNORECASE)
        if rev_suffix_match:
            cleaned_dn = dn[:rev_suffix_match.start()]
            logger.warning(
                f"⚠️  D.N. CLEANUP: Removed REV suffix from D.N. "
                f"'{dn}' → '{cleaned_dn}'"
            )
            result_data['drawing_number'] = cleaned_dn

    # ─── RAFAEL OCR CHECK: PN differs from DN by ≤2 chars ───
    # Catches OCR errors like FT-16912017 vs FT-15912017 (1 char diff).
    # Sanity B is skipped for RAFAEL (_searched_for_pn_field=True), so
    # this dedicated check handles the case: PN not in filename, DN is.
    if is_rafael and result_data.get('part_number') and result_data.get('drawing_number'):
        pn_ocr = str(result_data['part_number']).strip()
        dn_ocr = str(result_data['drawing_number']).strip()
        if pn_ocr.upper() != dn_ocr.upper():
            pn_in_fn = check_value_in_filename(pn_ocr, filename)
            dn_in_fn = check_value_in_filename(dn_ocr, filename)

            # ─── CAT.NO detection: purely numeric PN in RAFAEL is likely CAT.NO ───
            # RAFAEL P.N. values are almost always alphanumeric (e.g., BBLE4352A, G87148A).
            # Purely numeric values (e.g., 510030054 from CAT.NO 5100.30.02-4) are suspicious.
            if not pn_in_fn and dn_in_fn:
                pn_alnum = re.sub(r'[^A-Za-z0-9]', '', pn_ocr)
                is_purely_numeric = pn_alnum.isdigit() and len(pn_alnum) >= 6
                if is_purely_numeric:
                    # Try to recover real P.N. from pdfplumber text
                    recovered_pn = None
                    if pdfplumber_text:
                        recovered = extract_pn_dn_from_text(pdfplumber_text)
                        candidate_pn = recovered.get('part_number', '')
                        if (candidate_pn
                            and candidate_pn.upper() != dn_ocr.upper()
                            and candidate_pn != pn_ocr
                            and not candidate_pn.isdigit()
                            and len(candidate_pn) >= 5):
                            recovered_pn = candidate_pn

                    if recovered_pn:
                        logger.warning(
                            f"⚠️  RAFAEL CAT.NO FIX: PN '{pn_ocr}' is purely numeric "
                            f"(likely CAT.NO) — recovered real P.N. from pdfplumber: '{recovered_pn}'"
                        )
                        result_data['part_number'] = recovered_pn
                    else:
                        logger.warning(
                            f"⚠️  RAFAEL CAT.NO FIX: PN '{pn_ocr}' is purely numeric "
                            f"(likely CAT.NO, not P.N.) — replacing with Drawing# '{dn_ocr}'"
                        )
                        result_data['part_number'] = dn_ocr
                else:
                    pn_n = re.sub(r'[^a-z0-9]', '', pn_ocr.lower())
                    dn_n = re.sub(r'[^a-z0-9]', '', dn_ocr.lower())
                    if len(pn_n) == len(dn_n):
                        ocr_diffs = sum(1 for a, b in zip(pn_n, dn_n) if a != b)
                    else:
                        ocr_diffs = abs(len(pn_n) - len(dn_n)) + \
                            sum(1 for a, b in zip(pn_n, dn_n) if a != b)
                    if ocr_diffs <= 2:
                        logger.warning(
                            f"⚠️  RAFAEL OCR FIX: PN '{pn_ocr}' differs from DN '{dn_ocr}' "
                            f"by only {ocr_diffs} char(s) — OCR error. Replacing PN with DN."
                        )
                        result_data['part_number'] = dn_ocr

    # ─── P.N. SANITY CHECK A ───
    if result_data.get('part_number') and result_data.get('drawing_number'):
        pn = str(result_data['part_number']).strip()
        dwg = str(result_data['drawing_number']).strip()
        digit_count = sum(1 for c in pn if c.isdigit())

        if digit_count < 2:
            dwg_digits = sum(1 for c in dwg if c.isdigit())
            if dwg_digits >= 2:
                logger.warning(
                    f"⚠️  P.N. SANITY A: '{pn}' has only {digit_count} digit(s) — "
                    f"Copying from Drawing# '{dwg}'"
                )
                result_data['part_number'] = dwg
            else:
                fn_pn = extract_part_number_from_filename(file_path)
                if fn_pn and len(re.findall(r'\d', fn_pn)) >= 2:
                    logger.warning(
                        f"⚠️  P.N. SANITY A: Both P.N.='{pn}' and Drawing#='{dwg}' have "
                        f"insufficient digits — Using filename: '{fn_pn}'"
                    )
                    result_data['part_number'] = fn_pn
                    result_data['drawing_number'] = fn_pn

    # ─── P.N. SANITY CHECK B: project/assembly P.N. ───
    # SKIP when _searched_for_pn_field=True — RAFAEL Stage 1 explicitly found
    # separate P.N. and Drawing# fields in the title block.
    # The difference is intentional (e.g., FTLS04013A vs 8H-A48192).
    if (result_data.get('part_number') and result_data.get('drawing_number')
            and not result_data.get('_searched_for_pn_field', False)):
        pn = str(result_data['part_number']).strip()
        dwg = str(result_data['drawing_number']).strip()

        if pn.upper() != dwg.upper():
            pn_in_filename = check_value_in_filename(pn, filename)
            dwg_in_filename = check_value_in_filename(dwg, filename)

            if not pn_in_filename and dwg_in_filename:
                logger.warning(
                    f"⚠️  P.N. SANITY B: '{pn}' not in filename, "
                    f"but Drawing# '{dwg}' is — likely project/assembly P.N. "
                    f"Replacing with Drawing#"
                )
                result_data['part_number'] = dwg

    # ─── P.N. SANITY CHECK C: truncated part number ───
    if result_data.get('part_number') and result_data.get('drawing_number'):
        pn = str(result_data['part_number']).strip()
        dwg = str(result_data['drawing_number']).strip()

        pn_norm = re.sub(r'[^A-Za-z0-9]', '', pn).upper()
        dwg_norm = re.sub(r'[^A-Za-z0-9]', '', dwg).upper()

        if (pn_norm != dwg_norm
            and len(pn_norm) >= 3
            and len(dwg_norm) > len(pn_norm)
            and pn_norm in dwg_norm):

            fn_norm = re.sub(r'[^A-Za-z0-9]', '', Path(file_path).stem).upper()
            dwg_in_filename = dwg_norm in fn_norm or any(
                dwg_norm == re.sub(r'[^A-Za-z0-9]', '', t).upper()
                for t in re.findall(r'[A-Za-z0-9]+', Path(file_path).stem)
                if len(t) > 4
            )

            if dwg_in_filename:
                logger.warning(
                    f"⚠️  P.N. SANITY C: '{pn}' is truncated substring of Drawing# '{dwg}' "
                    f"(Drawing# confirmed in filename) — replacing P.N. with Drawing#"
                )
                result_data['part_number'] = dwg
            else:
                logger.info(
                    f"ℹ️  P.N. SANITY C: '{pn}' is substring of Drawing# '{dwg}' "
                    f"but Drawing# not in filename — keeping P.N. as-is"
                )

    # ─── P.N. SANITY CHECK D: P.N. identical to Drawing# in RAFAEL ───
    pn_str_d = str(result_data.get('part_number', '')).strip()
    pn_in_filename_d = check_value_in_filename(pn_str_d, Path(file_path).stem)
    pn_has_digits_d = len(re.findall(r'\d', pn_str_d)) >= 3

    if (result_data.get('_searched_for_pn_field', False)
        and not (pn_in_filename_d and pn_has_digits_d)
        and result_data.get('part_number')
        and result_data.get('drawing_number')
        and str(result_data['part_number']).strip().upper() == str(result_data['drawing_number']).strip().upper()):

        current_pn = str(result_data['part_number']).strip()
        logger.warning(
            f"⚠️  P.N. SANITY D: P.N.='{current_pn}' == Drawing# → "
            f"Vision likely confused fields. Searching for real P.N...."
        )

        recovered_pn = None

        # Source 1: pdfplumber text
        if pdfplumber_text:
            pdf_retry = extract_pn_dn_from_text(pdfplumber_text)
            pdfplumber_pn_retry = pdf_retry.get('part_number', '')
            if (pdfplumber_pn_retry
                and pdfplumber_pn_retry.upper() != current_pn.upper()
                and len(pdfplumber_pn_retry) >= 5):
                recovered_pn = pdfplumber_pn_retry
                logger.info(f"  → Recovered P.N. from pdfplumber (dedup): '{recovered_pn}'")

        # Source 2: NOTES text pattern
        if not recovered_pn:
            notes_text = str(result_data.get('notes_full_text') or '')
            m = re.search(
                r'(?:PART\s*NUMBER|P\.?N\.?)\s*(?:FROM\s+\w+\s+TO\s+)?([A-Za-z0-9][\w\-]{5,30})',
                notes_text, re.IGNORECASE
            )
            if m:
                candidate = m.group(1).strip()
                if candidate.upper() != current_pn.upper():
                    recovered_pn = candidate
                    logger.info(f"  → Recovered P.N. from NOTES: '{recovered_pn}'")

        if recovered_pn:
            logger.warning(
                f"⚠️  P.N. SANITY D: Replacing '{current_pn}' with recovered P.N. '{recovered_pn}'"
            )
            result_data['part_number'] = recovered_pn
        else:
            logger.info(f"  → Could not recover real P.N. — keeping '{current_pn}'")

    # ─── D.N. SANITY CHECK E: Drawing# not in filename but P.N. is ───
    # For RAFAEL: PN ≠ DN is normal (PN=L54095A, DN=R00-263018).
    # Only replace DN with PN if DN looks invalid (too short, NOTE prefix, etc.)
    if result_data.get('part_number') and result_data.get('drawing_number'):
        pn_e = str(result_data['part_number']).strip()
        dwg_e = str(result_data['drawing_number']).strip()
        if pn_e.upper() != dwg_e.upper():
            pn_in_fn_e = check_value_in_filename(pn_e, filename)
            dwg_in_fn_e = check_value_in_filename(dwg_e, filename)
            if pn_in_fn_e and not dwg_in_fn_e:
                # Check if DN looks like a valid number (not a NOTE line, etc.)
                dn_looks_valid = (
                    len(dwg_e) >= 5
                    and any(c.isdigit() for c in dwg_e)
                    and not re.match(r'^\d+\.', dwg_e)  # skip NOTE references like "4.TOP"
                )

                if is_rafael and dn_looks_valid:
                    # RAFAEL: Check similarity — OCR error or legitimate PN≠DN?
                    pn_norm_e = re.sub(r'[^a-z0-9]', '', pn_e.lower())
                    dn_norm_e = re.sub(r'[^a-z0-9]', '', dwg_e.lower())

                    if len(pn_norm_e) == len(dn_norm_e):
                        char_diffs = sum(1 for a, b in zip(pn_norm_e, dn_norm_e) if a != b)
                    else:
                        char_diffs = abs(len(pn_norm_e) - len(dn_norm_e)) + \
                            sum(1 for a, b in zip(pn_norm_e, dn_norm_e) if a != b)

                    if char_diffs <= 2:
                        # Very similar — likely OCR error, DN is in filename so it's correct
                        logger.warning(
                            f"⚠️  RAFAEL OCR FIX: PN '{pn_e}' differs from DN '{dwg_e}' "
                            f"by only {char_diffs} char(s) — likely OCR error. "
                            f"Replacing PN with DN."
                        )
                        result_data['part_number'] = dwg_e
                    else:
                        # Very different — legitimate PN≠DN (e.g., L54095A vs R00-263018)
                        logger.info(
                            f"ℹ️  RAFAEL: PN '{pn_e}' differs from DN '{dwg_e}' "
                            f"by {char_diffs} chars — keeping both (legitimate PN≠DN)"
                        )
                else:
                    # Non-RAFAEL, or DN looks wrong → replace
                    logger.warning(
                        f"⚠️  D.N. SANITY E: Drawing# '{dwg_e}' not in filename, "
                        f"but P.N. '{pn_e}' is — replacing Drawing# with P.N."
                    )
                    result_data['drawing_number'] = pn_e

    # ─── Fallback: copy between P.N. and Drawing# ───
    searched_for_pn = result_data.get('_searched_for_pn_field', False)

    if not result_data.get('part_number') and result_data.get('drawing_number'):
        drawing_in_filename = check_value_in_filename(result_data['drawing_number'], filename)
        if not searched_for_pn or drawing_in_filename or not is_cage_code(result_data['drawing_number']):
            if not searched_for_pn:
                result_data['part_number'] = result_data['drawing_number']
                logger.info(f"Part# copied from Drawing# (no explicit P.N. field found)")
            else:
                logger.info(f"P.N. field was searched explicitly - not using Drawing# as fallback")
        else:
            logger.info(f"Cannot copy Drawing# to Part# - it's a CAGE CODE!")

    elif not result_data.get('drawing_number') and result_data.get('part_number'):
        if not is_cage_code(result_data['part_number']):
            result_data['drawing_number'] = result_data['part_number']
            logger.info(f"Drawing# copied from Part#")
        else:
            logger.info(f"Cannot copy Part# to Drawing# - it's a CAGE CODE!")

    # ─── FREE FALLBACKS: PL Part Number → Filename extraction ───
    if not result_data.get('part_number'):
        fallback_pn = None
        fallback_source = None

        # Priority 1: PL main part number
        pl_pn = result_data.get('pl_main_part_number', '')
        if pl_pn and pl_pn != 'MULTIPLE':
            fallback_pn = pl_pn
            fallback_source = 'PL Part Number'

        # Priority 2: Extract from filename
        if not fallback_pn:
            fn_pn = extract_part_number_from_filename(file_path)
            if fn_pn:
                fallback_pn = fn_pn
                fallback_source = 'filename extraction'

        if fallback_pn:
            result_data['part_number'] = fallback_pn
            result_data['_fallback_source'] = fallback_source
            if not result_data.get('drawing_number'):
                result_data['drawing_number'] = fallback_pn
                logger.info(f"🆓 FREE FALLBACK: Part#/Drawing# set to '{fallback_pn}' (source: {fallback_source})")
            else:
                logger.info(f"🆓 FREE FALLBACK: Part# set to '{fallback_pn}' (source: {fallback_source})")
        else:
            if not result_data.get('drawing_number'):
                logger.warning(f"⚠️  No part_number or drawing_number could be determined (all fallbacks exhausted)")
            else:
                logger.warning(f"⚠️  No part_number could be determined (all fallbacks exhausted, drawing# exists: '{result_data['drawing_number']}')")

    # ─── IAI rule: Part Number and Drawing Number must be identical ───
    if is_iai:
        part_num = str(result_data.get('part_number') or '').strip()
        drawing_num = str(result_data.get('drawing_number') or '').strip()

        def _norm_id(value: str) -> str:
            return re.sub(r'[^A-Za-z0-9]', '', str(value or '')).upper()

        if part_num and not drawing_num:
            result_data['drawing_number'] = part_num
            logger.info(f"IAI rule: Drawing# set from Part#: {part_num}")
        elif drawing_num and not part_num:
            result_data['part_number'] = drawing_num
            logger.info(f"IAI rule: Part# set from Drawing#: {drawing_num}")
        elif part_num and drawing_num and _norm_id(part_num) != _norm_id(drawing_num):
            part_in_filename = check_value_in_filename(part_num, filename)
            drawing_in_filename = check_value_in_filename(drawing_num, filename)

            if part_in_filename and not drawing_in_filename:
                chosen = part_num
                reason = "part matched filename"
            elif drawing_in_filename and not part_in_filename:
                chosen = drawing_num
                reason = "drawing matched filename"
            else:
                chosen = part_num if len(_norm_id(part_num)) >= len(_norm_id(drawing_num)) else drawing_num
                reason = "best OCR candidate"

            result_data['part_number'] = chosen
            result_data['drawing_number'] = chosen
            logger.info(f"IAI rule: unified Part#/Drawing# to '{chosen}' ({reason})")

    return result_data


def calculate_confidence(
    result_data: dict,
    filename: str,
    file_path: str,
) -> dict:
    """
    Calculate confidence level (full/high/medium/low) based on filename matching.
    Updates result_data in place and returns it.
    """
    # Calculate if numbers are in filename (partial match)
    part_in_filename = check_value_in_filename(result_data.get('part_number'), filename)
    drawing_in_filename = check_value_in_filename(result_data.get('drawing_number'), filename)

    # Check for exact match
    part_exact_match = check_exact_match_in_filename(result_data.get('part_number'), filename)
    drawing_exact_match = check_exact_match_in_filename(result_data.get('drawing_number'), filename)

    result_data['part_in_filename'] = part_in_filename
    result_data['drawing_in_filename'] = drawing_in_filename

    # Calculate confidence level based on filename matching
    if part_exact_match and drawing_exact_match:
        confidence_level = 'full'
        needs_review = ""
        logger.info(f"Confidence: FULL (exact match for both part# and drawing# in filename)")
    elif part_in_filename:
        confidence_level = 'high'
        needs_review = ""
        if drawing_in_filename:
            logger.info(f"Confidence: HIGH (part# in filename, also drawing#)")
        else:
            logger.info(f"Confidence: HIGH (part# in filename)")
    elif drawing_in_filename:
        confidence_level = 'medium'
        needs_review = " בדיקה"
        logger.info(f"Confidence: MEDIUM (only drawing# in filename, no part#)")
    else:
        confidence_level = 'low'
        needs_review = " בעייתי"
        logger.info(f"Confidence: LOW (neither part# nor drawing# in filename)")

    result_data['confidence_level'] = confidence_level
    result_data['needs_review'] = needs_review

    # Update validation_warnings with confidence info
    existing_warnings = result_data.get('validation_warnings', '')
    if confidence_level == 'low':
        conf_warning = "Low confidence - numbers not in filename"
    elif confidence_level == 'medium':
        conf_warning = "Medium confidence - only one number in filename"
    elif confidence_level == 'high':
        conf_warning = "High confidence - both numbers partially in filename"
    else:  # full
        conf_warning = ""

    if conf_warning:
        if existing_warnings:
            result_data['validation_warnings'] = f"{existing_warnings}; {conf_warning}"
        else:
            result_data['validation_warnings'] = conf_warning

    return result_data
