"""
P.N. and Drawing# extraction and voting logic.
================================================

Extracted from customer_extractor_v3_dual.py for modularity.

Functions:
- deduplicate_line(): Fix doubled characters from pdfplumber extraction
- extract_pn_dn_from_text(): Extract P.N. and Drawing# from raw text
- vote_best_pn(): Vote between 3 extraction methods for best P.N.
"""

import re
from typing import Tuple

from src.utils.logger import get_logger

logger = get_logger(__name__)


def deduplicate_line(line: str) -> str:
    """
    Remove doubled characters from pdfplumber extraction.
    Some PDFs produce "MMMMAA7766JJ880000220033BB" instead of "MMA76J800203B"
    """
    s = line.strip()
    if len(s) < 6 or len(s) % 2 != 0:
        return line
    for i in range(0, len(s) - 1, 2):
        if s[i] != s[i + 1]:
            return line  # Not fully doubled
    return s[::2]  # Take every other char


def extract_pn_dn_from_text(text: str) -> dict:
    """
    Extract P.N. and DRAWING NO. from raw text (pdfplumber or Tesseract).
    Works on vector PDFs with near-perfect accuracy.

    Returns: {'part_number': '...', 'drawing_number': '...'}
    """
    result = {'part_number': '', 'drawing_number': ''}

    if not text:
        return result

    # ─── Step 0: Fix doubled characters (common in some RAFAEL PDFs) ───
    lines = text.split('\n')
    fixed_lines = [deduplicate_line(l) for l in lines]
    text = '\n'.join(fixed_lines)

    # --- P.N. extraction ---
    pn_patterns = [
        # "P.N. FTLS04009A" or "P.N.  BO44666A"
        r'P\.N\.?\s+([A-Za-z0-9][\w\-\.]{3,30})',
        # "PN  BO44666A"
        r'\bPN\s+([A-Za-z0-9][\w\-\.]{3,30})',
        # "PART NO. XXXX" or "PART NUMBER XXXX"
        r'PART\s*(?:NO|NUMBER|NUM)\.?\s+([A-Za-z0-9][\w\-\.]{3,30})',
    ]

    skip_values = {'SHT', 'OF', 'SEE', 'EDR', 'TITLE', 'SIZE', 'REV', 'SCALE',
                   'CLASS', 'CAGE', 'CODE', 'DATE', 'SHEET', 'DWG', 'DRAWING',
                   # Common English words that appear near "PART NUMBER" in NOTES:
                   'FROM', 'TO', 'CHANGED', 'REPLACED', 'WAS', 'THE', 'FOR',
                   'WITH', 'AND', 'NOT', 'PER', 'ACC', 'NONE', 'NULL', 'ITEM',
                   'DESCRIPTION', 'MATERIAL', 'QUANTITY', 'UNIT', 'EACH',
                   # Title block field labels
                   'DESIGN', 'DESIGNER', 'DESIGNED',
                   }

    # ─── Context filter: skip P.N. references in assembly/installation notes ───
    # e.g. "P.N 315060104 ON PIN Iso 8734" — this is a component P.N., not the main part
    _NOTE_CONTEXT_AFTER_RE = re.compile(
        r'\s+(?:'
        r'ON\s+(?:PIN|SHAFT|BOLT|SCREW|NUT|BUSHING|LINK|WASHER|SPRING|PLATE|RING|PLUG|PART|COMPONENT)'
        r'|OR\s+(?:PIN|DIN|ISO|BOLT|SCREW)'
        r'|(?:ISO|DIN)\s+\d'
        r'|(?:Iso|iso)\s+\d'
        r')',
        re.IGNORECASE
    )
    # Line-level filter: skip if the line containing P.N. has hardware/assembly keywords
    _NOTES_LINE_RE = re.compile(
        r'(?:APPLY|INSERT|VERIFY|INSPECT|LOCTITE|ASSEMBLE|TORQUE'
        r'|DIN\s+\d|ISO\s+\d{3,}|Iso\s+\d{3,})',
        re.IGNORECASE
    )

    def _is_note_context(match_obj, src_text):
        """Check if a P.N. match is in assembly/inspection notes context."""
        # Check text AFTER the match
        after_text = src_text[match_obj.end():match_obj.end() + 50]
        if _NOTE_CONTEXT_AFTER_RE.match(after_text):
            return True
        # Check the full LINE containing this P.N. for hardware/assembly keywords
        line_start = src_text.rfind('\n', 0, match_obj.start()) + 1
        line_end = src_text.find('\n', match_obj.end())
        if line_end == -1:
            line_end = len(src_text)
        full_line = src_text[line_start:line_end]
        if _NOTES_LINE_RE.search(full_line):
            return True
        return False

    # ─── PRIORITY 1: Multi-line title-block P.N. (header then value below) ───
    # Title block format: "P.N.  SHT  OF\n<value on next 1-3 lines>"
    # Search from END of text (title block is at the bottom of the drawing)
    pn_header_matches = list(re.finditer(
        r'(?:P\.N\.?|PN)\s*(?:SHT\s*(?:OF)?)?\s*\n',
        text, re.IGNORECASE
    ))
    for m in reversed(pn_header_matches):
        # Search next 3 lines for best P.N. candidate
        rest = text[m.end():]
        next_lines = rest.split('\n')[:3]
        all_tokens = []
        for line in next_lines:
            raw_tokens = re.findall(r'[A-Za-z0-9][\w\-\.]{3,30}', line)
            # Apply dedup to each token (some PDFs double every character)
            all_tokens.extend(deduplicate_line(t) for t in raw_tokens)
        # Pick best: prefer mixed alpha+digit with length >= 6 (typical P.N.)
        best = None
        for tok in all_tokens:
            tok_clean = tok.rstrip('.,:;!?').upper()
            if tok_clean in skip_values or len(tok) < 5:
                continue
            has_digit = bool(re.search(r'[0-9]', tok))
            has_alpha = bool(re.search(r'[A-Za-z]', tok))
            if has_digit and has_alpha and len(tok) >= 6:
                if best is None or len(tok) > len(best):
                    best = tok
        # If no mixed found, try pure numeric >= 6 digits
        if not best:
            for tok in all_tokens:
                tok_clean = tok.rstrip('.,:;!?').upper()
                if tok_clean in skip_values:
                    continue
                if tok.isdigit() and len(tok) >= 6:
                    if best is None or len(tok) > len(best):
                        best = tok
        if best:
            result['part_number'] = best
            logger.debug(f"P.N. extracted via multi-line title block: '{best}'")
            break

    # ─── PRIORITY 2: Inline P.N. with context filtering ───
    if not result['part_number']:
        for pat in pn_patterns:
            for m in re.finditer(pat, text, re.IGNORECASE):
                candidate = m.group(1).strip()
                # Strip trailing punctuation before comparing to skip_values
                candidate_clean = candidate.rstrip('.,:;!?').upper()
                if candidate_clean in skip_values or len(candidate) < 4:
                    continue
                # Check if this P.N. is a component reference in assembly notes
                if _is_note_context(m, text):
                    logger.debug(f"Skipped notes-context P.N.: '{candidate}'")
                    continue
                result['part_number'] = candidate
                break
            if result['part_number']:
                break

    # --- DRAWING NO. extraction ---
    dn_patterns = [
        r'DRAWING\s*NO\.?\s+([A-Za-z0-9][\w\-\.]{3,30})',
        r'DWG\s*NO\.?\s+([A-Za-z0-9][\w\-\.]{3,30})',
        r'DRAWING\s*NUMBER\s+([A-Za-z0-9][\w\-\.]{3,30})',
    ]

    skip_dn = {'SIZE', 'REV', 'SCALE', 'A0', 'A1', 'A2', 'A3', 'A4', 'CLASS',
               'TOMER', 'RAFAEL', 'IAI',
               # Common English words that appear on drawings but are NOT drawing numbers
               'TEXTURE', 'FINISH', 'SURFACE', 'MATERIAL', 'GENERAL', 'NOTES',
               'TOLERANCE', 'TOLERANCES', 'UNLESS', 'SPECIFIED', 'DIMENSIONS',
               'TITLE', 'DESCRIPTION', 'WEIGHT', 'SHEET', 'CAGE', 'CODE',
               'INTERPRET', 'PROPRIETARY', 'CONFIDENTIAL', 'CHECKED',
               'APPROVED', 'DRAWN', 'DATE', 'NAME',
               # Title block field labels that appear near DWG NO.
               'DESIGN', 'DESIGNER', 'DESIGNED',}

    for pat in dn_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            candidate = m.group(1).strip()
            candidate_clean = candidate.rstrip('.,:;!?').upper()
            if candidate_clean not in skip_dn and len(candidate) >= 4:
                result['drawing_number'] = candidate
                break

    # ─── Fallback: Multi-line Drawing# (value on line after "DRAWING NO.") ───
    if not result['drawing_number']:
        m = re.search(
            r'DRAWING\s*NO\.?\s*(?:SIZE|REV)?\s*\n\s*([A-Za-z0-9][\w\-\.]{4,30})',
            text, re.IGNORECASE
        )
        if m:
            candidate = m.group(1).strip()
            candidate_clean = candidate.rstrip('.,:;!?').upper()
            if candidate_clean not in skip_dn and len(candidate) >= 5:
                result['drawing_number'] = candidate
                logger.debug(f"Drawing# extracted via multi-line: '{candidate}'")

    # ─── Fallback: Multi-line DWG NO. with REV. on same line (TOMER format) ───
    if not result['drawing_number']:
        m = re.search(
            r'DWG\s*NO\.?\s*(?:REV\.?)?\s*\n\s*'
            r'(?:\d+\s+)?'
            r'(?:TOMER|RAFAEL|IAI)?\s*'
            r'([A-Za-z0-9][\w\-\.]{4,30})',
            text, re.IGNORECASE
        )
        if m:
            candidate = m.group(1).strip()
            candidate_clean = candidate.rstrip('.,:;!?').upper()
            if candidate_clean not in skip_dn and len(candidate) >= 5:
                result['drawing_number'] = candidate
                logger.debug(f"Drawing# extracted via multi-line DWG NO.: '{candidate}'")

    return result


def vote_best_pn(vision_pn: str, pdfplumber_pn: str, tesseract_pn: str,
                 filename: str) -> Tuple[str, str]:
    """
    Vote between 3 extraction methods.
    Returns: (best_pn, source_description)

    Priority logic:
    1. If 2+ methods agree -> use consensus
    2. If one matches filename -> prefer it
    3. If pdfplumber has result -> prefer it (most reliable for vector PDFs)
    4. Otherwise -> use Vision result
    """
    candidates = {
        'vision': vision_pn or '',
        'pdfplumber': pdfplumber_pn or '',
        'tesseract': tesseract_pn or '',
    }

    # Normalize for comparison
    def norm(s):
        return re.sub(r'[^A-Za-z0-9]', '', str(s)).upper()

    # Normalize filename tokens
    fn_tokens = set(re.findall(r'[A-Za-z0-9]+', filename.upper()))

    def matches_filename(val):
        n = norm(val)
        if not n or len(n) < 4:
            return False
        return any(n == t or n in t or t in n for t in fn_tokens if len(t) > 4)

    # ─── Reject pure-alpha common English words (not valid P.N. / Drawing#) ───
    _COMMON_ENGLISH_WORDS = {
        'TEXTURE', 'FINISH', 'SURFACE', 'MATERIAL', 'GENERAL', 'NOTES',
        'TOLERANCE', 'TOLERANCES', 'UNLESS', 'SPECIFIED', 'DIMENSIONS',
        'TITLE', 'DESCRIPTION', 'WEIGHT', 'SHEET', 'CAGE', 'CODE',
        'INTERPRET', 'PROPRIETARY', 'CONFIDENTIAL', 'CHECKED',
        'APPROVED', 'DRAWN', 'DATE', 'NAME', 'SCALE', 'REVISION',
        'BRACKET', 'HOUSING', 'COVER', 'PLATE', 'ASSEMBLY',
        'SECTION', 'DETAIL', 'VIEW', 'ITEM', 'PART', 'NUMBER',
        'DESIGN', 'DESIGNER', 'DESIGNED',
    }

    def _is_common_word(val):
        """Check if value is a pure-alpha common English word (not a valid ID)."""
        n = norm(val)
        if not n:
            return False
        alpha_only = re.sub(r'[0-9]', '', n)
        return alpha_only == n and n in _COMMON_ENGLISH_WORDS

    # Strategy 1: Consensus (2+ agree) — but reject common English words
    #   Also reject consensus if it doesn't match filename but another candidate does.
    normed = {k: norm(v) for k, v in candidates.items() if v and v != 'N/A'}
    for src1, n1 in normed.items():
        for src2, n2 in normed.items():
            if src1 != src2 and n1 == n2:
                if _is_common_word(candidates[src1]):
                    logger.debug(f"🗳️ Consensus '{candidates[src1]}' rejected — common English word")
                    break
                # Check: if consensus doesn't match filename but a dissenting source does → reject consensus
                consensus_val = candidates[src1]
                if not matches_filename(consensus_val):
                    dissenting = [
                        s for s in candidates
                        if s != src1 and s != src2
                        and candidates[s] and candidates[s] != 'N/A'
                        and matches_filename(candidates[s])
                    ]
                    if dissenting:
                        logger.info(
                            f"🗳️ Consensus '{consensus_val}' rejected — "
                            f"doesn't match filename, but {dissenting[0]}='{candidates[dissenting[0]]}' does"
                        )
                        break
                return candidates[src1], f"consensus ({src1}+{src2})"

    # Strategy 2: Filename match
    for src in ['pdfplumber', 'vision', 'tesseract']:
        val = candidates[src]
        if val and val != 'N/A' and matches_filename(val):
            return val, f"filename match ({src})"

    # ─── Strategy 2.5: Filename extraction when no candidate matches ───
    # If NO candidate matched the filename, but the filename itself contains
    # a clear part number token → use it directly.
    # This catches cases where Vision returns garbage (e.g., TD-595 from specs)
    # and pdfplumber has no labeled field, but the filename IS the part number.
    any_matched = any(
        matches_filename(v) for v in candidates.values()
        if v and v != 'N/A'
    )
    if not any_matched:
        from src.services.extraction.filename_utils import extract_part_number_from_filename
        fn_pn = extract_part_number_from_filename(filename)
        if fn_pn:
            fn_norm = norm(fn_pn)
            has_digits = bool(re.search(r'[0-9]', fn_pn))
            long_enough = len(fn_norm) >= 5
            candidates_look_valid = any(
                len(norm(v)) >= 6 and bool(re.search(r'[0-9]', v))
                for v in candidates.values()
                if v and v != 'N/A'
            )
            if has_digits and long_enough and not candidates_look_valid:
                logger.info(
                    f"🗳️ Strategy 2.5: No candidate matched filename — "
                    f"using filename extraction: '{fn_pn}'"
                )
                return fn_pn, "filename extraction (no match)"

    # Strategy 3: Prefer pdfplumber (most reliable for vector)
    # BUT: validate it's actually a part number, not a random English word
    if pdfplumber_pn and pdfplumber_pn != 'N/A' and len(pdfplumber_pn) >= 4:
        # Strip dots/punctuation before checking if it's purely alphabetic
        pn_alpha_only = re.sub(r'[^A-Za-z0-9]', '', pdfplumber_pn)
        pn_upper = pn_alpha_only.upper()
        is_common_word = (
            pn_alpha_only.isalpha()
            and len(pn_alpha_only) <= 12
            and not re.search(r'[0-9]', pdfplumber_pn)
        )
        # Also reject common title-block field names (e.g. "Eng.Mgr")
        _TITLE_BLOCK_WORDS = {
            'ENGMGR', 'ENGINEER', 'MANAGER', 'CHECKER', 'APPROVED',
            'DRAWN', 'CHECKED', 'DATE', 'TITLE', 'DESCRIPTION',
            'REVISION', 'SCALE', 'MATERIAL', 'FINISH', 'WEIGHT',
            'UNLESS', 'TOLERANCES', 'INTERPRET', 'DIMENSIONS',
            'SHEET', 'SIZE', 'CAGE', 'CODE', 'PROPRIETARY',
            'DRAWNBY', 'CHECKEDBY', 'APPROVEDBY', 'ENGMANAGER',
            'DESIGN', 'DESIGNER', 'DESIGNED', 'DESIGNEDBY',
        }
        is_title_block = pn_upper in _TITLE_BLOCK_WORDS
        if not is_common_word and not is_title_block:
            return pdfplumber_pn, "pdfplumber (preferred)"
        else:
            reason = "title-block field" if is_title_block else "English word"
            logger.debug(f"🗳️ Rejected pdfplumber '{pdfplumber_pn}' — looks like {reason}, not P.N.")

    # Strategy 4: Vision default
    if vision_pn and vision_pn != 'N/A':
        return vision_pn, "vision (default)"

    # Strategy 5: Tesseract last resort
    if tesseract_pn and tesseract_pn != 'N/A':
        return tesseract_pn, "tesseract (last resort)"

    return vision_pn or '', "none matched"
