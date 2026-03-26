"""
Filename-based utilities for DrawingAI Pro.
============================================

Pure helper functions that compare / correct OCR-extracted values against
the original filename.  Moved from customer_extractor_v3_dual.py.
"""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# =========================================================================
# Known part-number patterns for filename extraction
# =========================================================================
# Each pattern is a compiled regex that matches a typical part-number token
# inside an underscore-delimited filename such as:
#   33956_MA-S9160-1000_SHEET1_04112025.pdf
#   DWG_12345-678-901_REV_A.pdf
#   MA-S9160-1000.pdf
#
# Tokens that are purely numeric dates (ddmmyyyy / yyyymmdd), sheet markers,
# revision markers, or short serial IDs are *excluded*.

_FILENAME_SKIP_PATTERNS: list[re.Pattern] = [
    # Sheet-like tokens: SHEET1, Sheet_2, SH3 …
    re.compile(r'^(?:SHEET|SH)\d*$', re.IGNORECASE),
    # Revision tokens: REV, REVA, REV_A, REV-01
    re.compile(r'^REV[_\-]?[A-Z0-9]*$', re.IGNORECASE),
    # Pure date tokens: 8 digits that look like ddmmyyyy or yyyymmdd
    re.compile(r'^\d{8}$'),
    # Purely numeric short IDs (≤4 digits, likely page/internal serial)
    # Note: 5+ digit numbers (e.g. RAFAEL drawing numbers) are allowed
    re.compile(r'^\d{1,4}$'),
    # Single letter tokens (e.g. "A", "B")
    re.compile(r'^[A-Za-z]$'),
    # Common filler words
    re.compile(r'^(?:DWG|DRAWING|PAGE|SCAN|COPY)$', re.IGNORECASE),
]

from src.core.constants import debug_print


# =========================================================================
# Simple filename checks
# =========================================================================

def check_value_in_filename(value: Optional[str], filename: str) -> bool:
    """בדיקה האם ערך מופיע בשם הקובץ"""
    if not value or value == "null":
        return False
    
    def normalize(s):
        if not s:
            return ""
        s = str(s).lower()
        s = re.sub(r'[^a-z0-9]', '', s)
        return s
    
    normalized_value = normalize(value)
    normalized_filename = normalize(filename)
    
    if len(normalized_value) < 3:
        return False
    
    return normalized_value in normalized_filename


def check_exact_match_in_filename(value: Optional[str], filename: str) -> bool:
    """בדיקה האם יש התאמה מלאה (exact match) בשם הקובץ"""
    if not value or value == "null":
        return False
    
    def normalize(s):
        if not s:
            return ""
        s = str(s).lower()
        s = re.sub(r'[^a-z0-9]', '', s)
        return s
    
    normalized_value = normalize(value)
    normalized_filename = normalize(filename)
    
    if len(normalized_value) < 3:
        return False
    
    # Check if the normalized value is exactly equal to the normalized filename
    if normalized_value == normalized_filename:
        return True
    
    # Check with boundaries using pattern matching
    # Pattern: value appears with word boundaries (start/end or separated by non-alphanumeric)
    pattern = r'(?:^|[^a-z0-9])' + re.escape(normalized_value) + r'(?:[^a-z0-9]|$)'
    return bool(re.search(pattern, normalized_filename))


def fix_zero_o_from_filename(value: Optional[str], filename: str) -> Optional[str]:
    """
    Align 0/O ambiguities to the filename token.
    Compares normalized (no separators) strings, and if only 0/O differ, returns corrected value.
    """
    if not value:
        return value
    
    # Normalize value (remove non-alphanumeric but preserve structure for reconstruction)
    val_orig = str(value).strip()
    val_normalized = re.sub(r'[^A-Za-z0-9]', '', val_orig).upper()
    
    if not val_normalized:
        return value
    
    # Extract all alphanumeric tokens from filename
    tokens = re.findall(r"[A-Za-z0-9]+", filename)
    
    for tok in tokens:
        tok_upper = tok.upper()
        
        # Must have same length after normalization
        if len(tok_upper) != len(val_normalized):
            continue
        
        # Check if all differences are only 0 vs O
        diff_ok = True
        has_ambig = False
        for a, b in zip(val_normalized, tok_upper):
            if a == b:
                continue
            # Allow only 0 <-> O substitutions
            if {a, b} <= {"0", "O"}:
                has_ambig = True
                continue
            # Any other difference = not a match
            diff_ok = False
            break
        
        if diff_ok and has_ambig:
            # Found a matching token - reconstruct value with corrected chars
            # Strategy: replace chars in val_orig at positions where 0/O differ
            corrected = list(val_orig)
            val_idx = 0
            for i, ch in enumerate(val_orig):
                if ch.isalnum():
                    if val_idx < len(tok_upper):
                        # If there's a 0/O difference, use filename version
                        if val_normalized[val_idx].upper() != tok_upper[val_idx]:
                            if {val_normalized[val_idx].upper(), tok_upper[val_idx]} <= {"0", "O"}:
                                corrected[i] = tok[val_idx]  # Use original case from token
                        val_idx += 1
            return ''.join(corrected)
    
    return value


# =========================================================================
# Smart OCR disambiguation system
# =========================================================================

# Character confusion matrix for OCR ambiguities
CHAR_CONFUSION_MATRIX = {
    'D': {'0': 0.85, 'O': 0.80},  # D דומה מאוד ל-0 ול-O
    'O': {'0': 0.95, 'D': 0.75},  # O כמעט לא ניתן להבחין מ-0
    '0': {'O': 0.95, 'D': 0.75},  # 0 (ספרה) דומה ל-O,D
    'I': {'1': 0.90, 'l': 0.85},  # I (אות) דומה ל-1 ול-l
    'l': {'1': 0.85, 'I': 0.80},  # l (אות קטנה) דומה ל-1 ו-I
    '1': {'I': 0.88, 'l': 0.85},  # 1 (ספרה) דומה ל-I ול-l
}


def _get_confusion_positions(value_str: str) -> Dict[int, list]:
    """
    זהה positions בstring שבהן יש ambiguity (D/0/O, I/1/l)
    
    Returns:
        {position: [list of possible char alternatives]}
    """
    positions = {}
    for i, ch in enumerate(value_str.upper()):
        if ch in CHAR_CONFUSION_MATRIX:
            positions[i] = list(CHAR_CONFUSION_MATRIX[ch].keys())
    return positions


def _generate_candidates(value_str: str, max_candidates: int = 10) -> List[Tuple[str, float]]:
    """
    Generate candidate alternatives for ambiguous characters.
    PRIORITY: O/0 confusion (most common), then other D/I/l/1 confusions.
    
    Returns:
        List of (candidate_string, confidence_score_between_0_and_1)
    """
    candidates = [(value_str, 1.0)]  # Original with confidence 100%
    
    if not value_str or len(value_str) < 3:
        return candidates
    
    confusion_positions = _get_confusion_positions(value_str)
    if not confusion_positions:
        return candidates
    
    # Strategy: Handle O/0 confusion separately (it's most common and tricky)
    # Add ALL O/0 substitutions, then add other confusions
    
    value_upper = value_str.upper()
    
    # Step 1: Generate O/0 candidates (ALL positions, not just top 2)
    oo_candidates = []
    for i, ch in enumerate(value_upper):
        if ch in ('O', '0'):  # Found an ambiguous char
            alternatives = ['0'] if ch == 'O' else ['O']
            for alt_ch in alternatives:
                alt_str = value_str[:i] + alt_ch + value_str[i+1:]
                confidence = 0.95 * 0.85  # O/0 similarity is 0.95
                oo_candidates.append((alt_str, confidence))
    
    # Add O/0 candidates (they're most reliable)
    candidates.extend(oo_candidates[:5])  # Limit to avoid explosion
    
    # Step 2: Generate OTHER confusions (D/I/l/1) for top 2 positions
    sorted_positions = sorted(
        [p for p in confusion_positions.keys() if value_upper[p] not in ('O', '0')]
    )
    
    for pos in sorted_positions[:2]:  # Top 2 non-O/0 confusions
        ch = value_str[pos].upper()
        if ch in CHAR_CONFUSION_MATRIX:
            alternatives = CHAR_CONFUSION_MATRIX[ch]
            
            for alt_ch, similarity in alternatives.items():
                alt_str = value_str[:pos] + alt_ch + value_str[pos+1:]
                confidence = similarity * 0.85
                candidates.append((alt_str, confidence))
    
    # Remove duplicates and sort by confidence
    seen = set()
    unique_candidates = []
    for cand, conf in candidates:
        if cand.upper() not in seen:
            seen.add(cand.upper())
            unique_candidates.append((cand, conf))
    
    # Sort by confidence descending, limit to max_candidates
    unique_candidates = sorted(unique_candidates, key=lambda x: x[1], reverse=True)[:max_candidates]
    
    return unique_candidates


def _score_candidate_against_filename(candidate: str, filename: str) -> float:
    """
    Score how well a candidate matches the filename.
    Higher score = better match.
    
    Scoring:
    - Exact match (after normalization): 100 points
    - Partial match (substring): 75 points
    - Length match: 10 points
    - Position match (starts at beginning): 5 points
    """
    if not candidate:
        return 0.0
    
    score = 0.0
    
    # Extract alphanumeric tokens from filename
    tokens = re.findall(r'[A-Za-z0-9]+', filename)
    cand_norm = re.sub(r'[^A-Za-z0-9]', '', candidate).upper()
    
    for token in tokens:
        token_norm = token.upper()
        
        # Exact match
        if cand_norm == token_norm:
            return 100.0
        
        # Partial match
        if cand_norm in token_norm or token_norm in cand_norm:
            score = max(score, 75.0)
        elif len(cand_norm) == len(token_norm):
            # Same length - measure similarity
            diff_count = sum(1 for a, b in zip(cand_norm, token_norm) if a != b)
            if diff_count <= 2:  # Only 1-2 chars different
                similarity = (len(cand_norm) - diff_count) / len(cand_norm)
                score = max(score, 60 + (similarity * 30))
    
    # Bonus for length match
    for token in tokens:
        if len(re.sub(r'[^A-Za-z0-9]', '', candidate)) == len(token):
            score += 10
    
    return score


def _disambiguate_part_number(ocr_value: str, filename: str) -> Dict[str, Any]:
    """
    Smart disambiguation of part number using filename as source of truth.
    Tries to fix D/0/O, I/1/l confusions.
    
    Returns:
        {
            'best_candidate': 'corrected_part_number',
            'confidence': float (0-1),
            'alternatives': [(candidate, score), ...],
            'method': 'exact_match' | 'filename_based' | 'smart_substitution'
        }
    """
    if not ocr_value:
        return {
            'best_candidate': ocr_value,
            'confidence': 0.0,
            'alternatives': [],
            'method': 'none'
        }
    
    ocr_value = str(ocr_value).strip()
    
    # Step 1: Generate candidates
    candidates = _generate_candidates(ocr_value)
    
    # Optional debug: log candidates for O/0 confusion cases
    cand_strs = [c for c, _ in candidates[:10]]
    if 'O' in ocr_value.upper() or '0' in ocr_value:
        debug_print(f"[DMBG] OCR: '{ocr_value}' | Candidates: {cand_strs} | Filename: '{filename}'")
    
    # Step 2: Score each candidate against filename
    scored_candidates = [
        (cand, _score_candidate_against_filename(cand, filename), conf)
        for cand, conf in candidates
    ]
    
    # Sort by filename score (descending)
    scored_candidates = sorted(scored_candidates, key=lambda x: x[1], reverse=True)
    
    if not scored_candidates:
        return {
            'best_candidate': ocr_value,
            'confidence': 0.5,
            'alternatives': [],
            'method': 'fallback'
        }
    
    best_candidate, filename_score, ocr_confidence = scored_candidates[0]
    
    # Calculate final confidence
    # If filename score is high, we trust the match
    if filename_score > 80:
        final_confidence = min(1.0, ocr_confidence * 1.2)  # Boost confidence if filename confirms
        method = 'exact_match' if filename_score >= 100 else 'filename_based'
    elif filename_score > 50:
        final_confidence = ocr_confidence * 0.9  # Slight reduction if partial match
        method = 'filename_based'
    else:
        final_confidence = ocr_confidence * 0.6  # Low confidence if no filename match
        method = 'smart_substitution'
    
    return {
        'best_candidate': best_candidate,
        'confidence': final_confidence,
        'alternatives': [
            (cand, f_score) for cand, f_score, _ in scored_candidates[:5]
        ],
        'method': method
    }


# =========================================================================
# Item number extraction from filename
# =========================================================================

def _extract_item_number_from_filename(filename: str) -> str:
    """
    Extract item/part number from filename.
    Looks for sequences of alphanumeric characters (especially those with letters and numbers).
    Examples: 3585410, 68A250781, etc.
    """
    if not filename:
        return ""
    
    # Remove extension and common suffixes
    name = Path(filename).stem.lower()
    
    # Remove common prefixes (MDMD_, MDRW_, etc.)
    prefixes_to_remove = ['mdmd_', 'mdrw_', 'md_', 'dwg_', 'drawing_', 'doc_', 'doc', 'pict', 'pict_', 'image_', 'photo_', 'render_']
    for prefix in prefixes_to_remove:
        if name.startswith(prefix):
            name = name[len(prefix):]
    
    # Remove common suffixes (PL, 3D, MODEL, etc.) and file type codes (_30, _25, _16, _99)
    suffixes_to_remove = ['_pl', '_3d', '_model', '_asm', '_assembly', 'pl', '3d', 'model', 'asm', '_30', '_25', '_16', '_99', '_15', '_24']
    for suffix in suffixes_to_remove:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
    
    # Extract tokens
    tokens = re.findall(r'[a-z0-9]+', name)
    
    # Skip very short tokens and common words
    skip_words = {'a', 'b', 'c', 'd', 'e', 'f', 'the', 'and', 'or', 'for', 'to', 'from'}
    valid_tokens = [t for t in tokens if len(t) >= 4 and t not in skip_words]
    
    # Prefer tokens with mix of letters and numbers, or pure numbers of reasonable length
    for token in valid_tokens:
        # Prefer tokens with mix of letters and numbers
        if any(c.isalpha() for c in token) and any(c.isdigit() for c in token):
            return token  # Mixed: prefer
    
    # Fallback: longest pure number token (5+ digits is likely a part number)
    for token in sorted(valid_tokens, key=len, reverse=True):
        if token.isdigit() and len(token) >= 5:
            return token  # Long number
    
    # Final fallback: longest valid token
    if valid_tokens:
        return max(valid_tokens, key=len)
    
    return ""


# =========================================================================
# Fuzzy matching helpers
# =========================================================================

def _normalize_item_number(item_num: str) -> str:
    """
    נרמול מספר פריט להשוואה:
    - מסיר סוגריים מרובעות ותוכן [D], [A], וכו'
    - מסיר סוגריים עגולות ותוכן
    - מסיר רווחים, קווים תחתונים, מקפים
    - מסיר נקודות, פסיקים וכל תווים מיוחדים אחרים
    - מסיר אפסים מיותרים בסוף (000, 00)
    - מחליף O/o ל-0 (תיקון בלבול OCR)
    - ממיר לאותיות קטנות (case-insensitive)
    
    דוגמאות:
    - "BO52931A [D]" -> "b052931a"
    - "B052931A-000" -> "b052931a"
    - "BO52931A.00" -> "b052931a"
    - "B-O52931A" -> "b052931a"
    - "Tube_Spacer" -> "tubespacer"
    - "PAL3804324" -> "pa13804324"    (L→1, עקבי בין PAL ל-pal)
    """
    if not item_num:
        return ""
    
    # הסר סוגריים מרובעות ותוכן
    normalized = re.sub(r'\[.*?\]', '', item_num)
    # הסר סוגריים עגולות ותוכן
    normalized = re.sub(r'\(.*?\)', '', normalized)
    # הסר רווחים, קווים תחתונים, מקפים
    normalized = re.sub(r'[\s_\-]', '', normalized)
    # הסר נקודות, פסיקים וכל תווים מיוחדים (חוץ מאותיות ומספרים)
    normalized = re.sub(r'[^\w]', '', normalized)
    # המר לאותיות קטנות — חשוב לעשות BEFORE תיקוני OCR כדי שהתוצאה תהיה עקבית
    # ללא תלות ב-case של הקלט (PAL vs pal נותנים אותו פלט)
    normalized = normalized.lower().strip()
    # החלף o ל-0 (תיקון בלבול OCR נפוץ בין אות o למספר 0)
    normalized = normalized.replace('o', '0')
    # החלף i/l ל-1 (תיקון בלבול OCR בין אות i, l למספר 1)
    normalized = normalized.replace('i', '1').replace('l', '1')
    
    # הסר אפסים מיותרים בסוף (000, 00, 0) - אבל לא אם זה כל המספר
    # חפש דפוס של 2 או יותר אפסים בסוף המחרוזת
    if len(normalized) > 3:
        normalized = re.sub(r'0+$', '', normalized)
    
    return normalized


def _fuzzy_char_equal(c1: str, c2: str) -> bool:
    """
    השוואת תווים עם עדינות לבלבול OCR
    - O == 0 (אות O == מספר 0)
    - I == 1 (אות I == מספר 1)
    - i == 1 (אות i קטנה == מספר 1)
    - l == 1 (אות l קטנה == מספר 1)
    
    לכל זוג אחר - חייב להיות בדיוק אותו דבר
    """
    if c1 == c2:
        return True
    
    # זוגות OCR confusion - עדינות שמפלטנו בפי OCR
    ocr_pairs = {('O', '0'), ('0', 'O'), ('o', '0'), ('0', 'o'),
                 ('I', '1'), ('1', 'I'), ('i', '1'), ('1', 'i'),
                 ('l', '1'), ('1', 'l')}
    
    if (c1, c2) in ocr_pairs:
        return True
    
    return False


def _fuzzy_substring_match(part_num: str, filename: str, min_length: int = 5) -> bool:
    """
    בדיקה אם part_num מופיע בתוך filename עם עדינות לבלבול OCR
    - חייב להיות לפחות min_length תווים מדוקט חוקיים
    - כטר תו צריך להיות או בדיוק כמו או אם OCR confusion
    - אם לא OCR confusion - לא היטבה
    
    דוגמאות:
    - part_num="b044666a", filename="bo44666a"  True (O==0 זוג OCR)
    - part_num="ui6543i", filename="il6543l"  True (I==1, l==1 זוגי OCR)
    - part_num="b044666a", filename="b155666a"  False (0!=1, לא זוג OCR)
    - part_num="b044666a", filename="b255666a"  False (0!=2, לא OCR confusion מוכר)
    """
    if len(part_num) < min_length or len(filename) < min_length:
        return False
    
    # חפש את part_num בתוך filename עם fuzzy matching
    for start_idx in range(len(filename) - len(part_num) + 1):
        match = True
        for i, c1 in enumerate(part_num):
            c2 = filename[start_idx + i]
            if not _fuzzy_char_equal(c1, c2):
                match = False
                break
        
        if match:
            return True
    
    return False


# =========================================================================
# Part-number extraction from filename (zero-cost fallback)
# =========================================================================

def extract_part_number_from_filename(file_path: str) -> Optional[str]:
    """
    Try to extract a likely part-number from the PDF filename.

    Common patterns handled (underscore / space / hyphen delimited):
        33956_MA-S9160-1000_SHEET1_04112025.pdf  →  MA-S9160-1000
        DWG_12345-678-901_REV_A.pdf              →  12345-678-901
        MA-S9160-1000.pdf                        →  MA-S9160-1000
        ABC-123_Sheet2.pdf                       →  ABC-123

    Strategy:
        1. Split the stem by underscores / spaces.
        2. Filter out tokens that are dates, sheet markers, revisions, etc.
        3. Among remaining tokens, prefer the longest one that contains both
           letters and digits (or is at least 5 chars with hyphens) — this is
           the most likely part number.  If all remaining tokens are purely
           numeric (>6 digits) keep the longest one.

    Returns None when no plausible part-number can be identified.
    """
    stem = Path(file_path).stem  # without extension

    if not stem or stem.startswith('.'):
        return None

    # Split by underscores and spaces (keep hyphens inside tokens)
    tokens = re.split(r'[_ ]+', stem)

    # ─── Fix: Split RAFAEL filename patterns ───
    # RAFAEL filenames: "BO27825A-A-PD-BO27825A_A.pdf"
    # Pattern: PN-REV-PD-PN (PD = Production Drawing)
    # Without this fix, "BO27825A-A-PD-BO27825A" stays as one token
    expanded_tokens = []
    for tok in tokens:
        # Split by -PD-, -MD-, -AS-, -DW-, -DR- (common RAFAEL separators)
        sub_parts = re.split(r'-(?:PD|MD|AS|DW|DR)-', tok, flags=re.IGNORECASE)
        if len(sub_parts) > 1:
            # Also split remaining by single-letter segments: "BO27825A-A" -> ["BO27825A", "A"]
            for part in sub_parts:
                expanded_tokens.extend(re.split(r'(?<=\w)-(?=[A-Z]$)', part))
        else:
            expanded_tokens.append(tok)
    tokens = expanded_tokens

    candidates: list[str] = []
    for tok in tokens:
        tok_stripped = tok.strip()
        if not tok_stripped:
            continue
        # Skip tokens matching any ignore pattern
        if any(pat.match(tok_stripped) for pat in _FILENAME_SKIP_PATTERNS):
            continue
        candidates.append(tok_stripped)

    if not candidates:
        return None

    def _score(token: str) -> int:
        """Higher is better.  Tokens with mixed alpha+digit content score highest."""
        has_alpha = bool(re.search(r'[A-Za-z]', token))
        has_digit = bool(re.search(r'[0-9]', token))
        length = len(token)
        mixed = 1000 if (has_alpha and has_digit) else 0
        return mixed + length

    best = max(candidates, key=_score)

    # Final sanity: reject very short tokens (< 3 chars) with no hyphens
    if len(best) < 3:
        return None

    return best
