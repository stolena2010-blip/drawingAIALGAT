"""
IAI-specific Drawing Extraction Stages
========================================
Israel Aerospace Industries - custom extraction logic

Extracted from customer_extractor_v3_dual.py
"""
import os
import re
import base64
import cv2
import numpy as np
from typing import Optional, Tuple, Dict
from openai import AzureOpenAI

from src.utils.logger import get_logger
from src.services.ai.vision_api import (
    _resolve_stage_call_config,
    _chat_create_with_token_compat,
    _call_vision_api_with_retry,
)
from src.services.extraction.ocr_engine import debug_print, MultiOCREngine, IAI_TOP_RED_FALLBACK_ENABLED
from src.services.extraction.stages_generic import (
    identify_drawing_layout,
    extract_basic_info,
    extract_processes_info,
    extract_notes_text,
    calculate_geometric_area,
)

logger = get_logger(__name__)


# 

def _extract_iai_top_red_identifier(img_bytes: bytes, ocr_engine: Optional[MultiOCREngine] = None) -> Dict[str, Optional[str]]:
    """
    Extract IAI top red identifier (drawing/part + revision) from a tiny upper area.
    We strongly enlarge this area because it is usually very small but highly reliable.
    """
    result = {
        'identifier': None,
        'revision': None,
        'raw_text': None
    }

    try:
        nparr = np.frombuffer(img_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        if img is None:
            return result

        h, w = img.shape[:2]
        if h < 50 or w < 50:
            return result

        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        mask1 = cv2.inRange(hsv, np.array([0, 50, 35]), np.array([14, 255, 255]))
        mask2 = cv2.inRange(hsv, np.array([165, 50, 35]), np.array([179, 255, 255]))
        red_mask = cv2.bitwise_or(mask1, mask2)
        red_mask = cv2.morphologyEx(red_mask, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8), iterations=1)

        contours, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        min_area = max(80, int(w * h * 0.00001))

        roi_candidates = []
        for cnt in contours:
            x, y, cw, ch = cv2.boundingRect(cnt)
            area = cw * ch
            if area < min_area or cw < 18 or ch < 6:
                continue

            x0 = max(0, x - max(14, int(cw * 0.35)))
            y0 = max(0, y - max(10, int(ch * 1.5)))
            x1 = min(w, x + cw + max(14, int(cw * 0.35)))
            y1 = min(h, y + ch + max(10, int(ch * 1.5)))

            score = float(area)
            if y < int(h * 0.35):
                score *= 1.18
            if x > int(w * 0.55):
                score *= 1.10
            if cw > ch * 2.0:
                score *= 1.15

            roi = img[y0:y1, x0:x1]
            if roi.size > 0:
                roi_candidates.append((score, roi))

        if not roi_candidates:
            fallback_h = int(h * 0.22)
            fallback_w = int(w * 0.60)
            roi_candidates = [
                (1.0, img[:fallback_h, :fallback_w]),
                (0.9, img[:fallback_h, w - fallback_w:]),
            ]

        roi_candidates = sorted(roi_candidates, key=lambda x: x[0], reverse=True)[:4]
        normalized_candidates = []

        try:
            import pytesseract
            tess_ok = True
            tess_path = os.getenv("TESSERACT_PATH")
            if tess_path and os.path.exists(tess_path):
                pytesseract.pytesseract.tesseract_cmd = tess_path
        except Exception:
            tess_ok = False

        MAX_ENLARGED_PIXELS = 4_000_000  # ~2000x2000 cap to avoid Tesseract slowdown
        for _, roi in roi_candidates:
            roi_h, roi_w = roi.shape[:2]
            desired_pixels = roi_h * roi_w * 36  # 6x each dim = 36x area
            if desired_pixels > MAX_ENLARGED_PIXELS:
                scale = (MAX_ENLARGED_PIXELS / (roi_h * roi_w)) ** 0.5
                scale = max(scale, 1.0)
            else:
                scale = 6.0
            enlarged = cv2.resize(roi, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
            b, g, r = cv2.split(enlarged)
            bg_max = cv2.max(b, g)
            red_only = cv2.subtract(r, bg_max)
            red_only = cv2.normalize(red_only, None, 0, 255, cv2.NORM_MINMAX)
            red_only = cv2.GaussianBlur(red_only, (3, 3), 0)
            _, bin_img = cv2.threshold(red_only, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            for angle in (0, 90):
                if angle == 0:
                    probe = bin_img
                elif angle == 90:
                    probe = cv2.rotate(bin_img, cv2.ROTATE_90_COUNTERCLOCKWISE)
                elif angle == 270:
                    probe = cv2.rotate(bin_img, cv2.ROTATE_90_CLOCKWISE)
                else:
                    probe = cv2.rotate(bin_img, cv2.ROTATE_180)

                if tess_ok:
                    try:
                        wl = "-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_/.:, "
                        t1 = pytesseract.image_to_string(probe, config=f"--oem 3 --psm 7 {wl}")
                        t2 = pytesseract.image_to_string(probe, config=f"--oem 3 --psm 6 {wl}")
                        for txt in (t1, t2):
                            if txt:
                                t = txt.upper().replace('\n', ' ').replace('\r', ' ')
                                t = re.sub(r'\s+', ' ', t).strip()
                                if t:
                                    normalized_candidates.append(t)
                    except Exception:
                        pass

                # Skip redundant ocr_engine.extract_all — direct Tesseract above is sufficient

        if not normalized_candidates:
            return result

        # IAI rule (per business logic):
        # - identifier appears AFTER RELEASED / release-related word
        # - revision appears AFTER explicit "REV:"
        drawing_pattern = re.compile(r'\b[A-Z0-9]{2,}-[A-Z0-9-]{2,}\b')
        rev_pattern = re.compile(r'\bREV(?:ISION)?\s*:\s*([A-Z0-9]{1,4})\b')
        release_kw_pattern = re.compile(r'\b(?:RELEASED|RELEASE|REL|ISSUED|ISSUE|APPROVED|REL\.?\s*DATE|RELEASE\s*DATE)\b')
        # Date-like tokens to exclude from identifier candidates (e.g., 16-DEC-24)
        date_like_pattern = re.compile(r'^\d{1,2}-[A-Z]{3}-\d{2,4}$')
        strong_id_pattern = re.compile(r'^[A-Z0-9]{2,}-[A-Z0-9]{2,}$')

        best_text = max(normalized_candidates, key=len)
        best_identifier = None
        best_revision = None
        best_score = -1

        for cand in normalized_candidates:
            m_rel = release_kw_pattern.search(cand)
            m_rev = rev_pattern.search(cand)

            # Revision strictly from REV:
            cand_revision = m_rev.group(1) if m_rev else None

            cand_identifier = None
            if m_rel:
                tail = cand[m_rel.end():]
                # Stop scan before REV part (if present)
                if m_rev:
                    tail = tail[:max(0, m_rev.start() - m_rel.end())]

                # Find identifier-like tokens in release tail
                token_candidates = re.findall(r'[A-Z0-9][A-Z0-9\-]{3,}', tail)
                for token in token_candidates:
                    if date_like_pattern.match(token):
                        continue
                    # Must look like real part/drawing identifier: includes hyphen and at least one digit
                    if not strong_id_pattern.match(token):
                        continue
                    if not re.search(r'\d', token):
                        continue
                    if '-' in token:
                        cand_identifier = token
                        break

            # Best match: has release-based identifier
            if cand_identifier:
                cand_score = 100 + (20 if cand_revision else 0)
                if cand_score > best_score:
                    best_score = cand_score
                    best_text = cand
                    best_identifier = cand_identifier
                    best_revision = cand_revision

        if not best_identifier:
            m_id = drawing_pattern.search(best_text)
            if m_id:
                candidate = m_id.group(0)
                if strong_id_pattern.match(candidate) and re.search(r'\d', candidate) and not date_like_pattern.match(candidate):
                    best_identifier = candidate

        if not best_revision:
            m_rev = rev_pattern.search(best_text)
            if m_rev:
                best_revision = m_rev.group(1)

        result['raw_text'] = best_text
        result['identifier'] = best_identifier
        result['revision'] = best_revision
        return result

    except Exception:
        return result

def identify_drawing_layout_iai(image_b64: str, client: AzureOpenAI) -> Tuple[Optional[str], int, int]:
    """
    שלב 0.5 ל-IAI: זיהוי סדר Title Block - גרסת IAI
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי לשנות את הפרומפט כאן
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return identify_drawing_layout(image_b64, client)


def extract_basic_info_iai(
    text_snippet: str,
    image_b64: str,
    client: AzureOpenAI,
    ocr_engine: Optional[MultiOCREngine] = None,
    use_top_red_fallback: bool = False
) -> Tuple[Optional[Dict], int, int]:
    """
    שלב 1 ל-IAI: חילוץ מידע בסיסי - מותאם לפורמט IAI
    משתמש ב-Title Block מוגדל לדיוק מקסימלי
    
     עבור IAI: P.N. ו-DRAWING NO. ממוקומים בתאים שונים בתוך Title Block
     צריך לחפש כל אחד במיקום הספציפי שלו
    """
    # Start from generic extraction
    data, tokens_in, tokens_out = extract_basic_info(text_snippet, image_b64, client)

    if data is None:
        data = {}

    # For IAI, use top-red extraction ONLY as fallback (after regular area attempts failed)
    if use_top_red_fallback:
        try:
            img_bytes = base64.b64decode(image_b64)
            top_info = _extract_iai_top_red_identifier(img_bytes, ocr_engine)
            top_identifier = (top_info.get('identifier') or '').strip()
            top_revision = (top_info.get('revision') or '').strip()

            if top_identifier:
                data['drawing_number'] = top_identifier
                data['part_number'] = top_identifier
                data['_iai_top_red_used'] = True
                data['_iai_top_red_raw_text'] = top_info.get('raw_text')
                logger.info(f"IAI top-red fallback override: drawing/part = {top_identifier}")

                if top_revision:
                    data['revision'] = top_revision
                    logger.info(f"IAI top-red fallback override: revision = {top_revision}")
        except Exception as e:
            logger.info(f"IAI top-red extraction skipped: {e}")

    # Force customer name for IAI path
    data['customer_name'] = 'IAI'

    return data, tokens_in, tokens_out


def extract_processes_info_iai(text_snippet: str, image_b64: str, client: AzureOpenAI, **kwargs) -> Tuple[Optional[Dict], int, int]:
    """
    שלב 2 ל-IAI: תהליכים ומפרטים - גרסת IAI
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי לשנות את הפרומפט כאן
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return extract_processes_info(text_snippet, image_b64, client, **kwargs)


def extract_notes_text_iai(text_snippet: str, image_b64: str, client: AzureOpenAI, **kwargs) -> Tuple[Optional[str], int, int]:
    """
    שלב 3 ל-IAI: הנחיות מלאות - גרסת IAI
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי לשנות את הפרומפט כאן
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return extract_notes_text(text_snippet, image_b64, client, **kwargs)


def extract_area_info_iai(text_snippet: str, image_b64: str, client: AzureOpenAI) -> Tuple[Optional[str], int, int]:
    """
    שלב 4 ל-IAI: חישוב שטח - גרסת IAI
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי לשנות את הפרומפט כאן
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return calculate_geometric_area(text_snippet, image_b64, client)
