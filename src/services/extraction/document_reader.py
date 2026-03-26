"""
Document Reader – email parsing, OCR, and order/quote item extraction
=====================================================================
Extracted from customer_extractor_v3_dual.py  (Phase 2.7)
"""

import base64
import io
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

import cv2
import numpy as np
import pandas as pd
import pdfplumber
from openai import AzureOpenAI

from src.core.constants import (
    RESPONSE_FORMAT,
    MODEL_RUNTIME,
    STAGE_EMAIL_QUANTITIES,
    STAGE_ORDER_ITEM_DETAILS,
    debug_print,
)
from src.services.ai.vision_api import (
    _call_vision_api_with_retry,
    _chat_create_with_token_compat,
    _resolve_stage_call_config,
)
from src.services.extraction.filename_utils import _normalize_item_number
from src.utils.logger import get_logger
from src.utils.prompt_loader import load_prompt

logger = get_logger(__name__)


# ── public API ──────────────────────────────────────────────────────────
__all__ = [
    "_read_email_content",
    "_extract_quantities_from_order_pdf",
    "_extract_text_via_ocr",
    "_extract_item_details_from_documents",
]


# ────────────────────────────────────────────────────────────────────────
def _read_email_content(folder_path: Path, client: Optional[AzureOpenAI] = None) -> Dict[str, Any]:
    """
    קריאת קובץ מייל מתקייה וחילוץ כמויות כלליות ותיאור עבודה כללי באמצעות AI
    
    Args:
        folder_path: התקייה לחיפוש קובץ המייל
        client: Azure OpenAI client (אופציונלי - אם None, לא יחלץ כמויות/תיאור)
        
    Returns:
        Dictionary עם 'subject', 'body', 'found', 'general_quantities', 'part_quantities',
        'work_description', 'part_work_descriptions', 'general_work_description', 'work_description_negation'
    """
    logger.debug(f"[EMAIL] Looking for email files in {folder_path.name}...")
    result = {
        'subject': '', 'from': '', 'body': '', 'found': False,
        'general_quantities': [], 'part_quantities': {},
        'work_description': '', 'quantity_summary': '',
        # ── work description fields (from AI) ──
        'part_work_descriptions': {},       # per-part work descriptions from email
        'general_work_description': '',     # general work description from email
        'work_description_negation': '',    # negation patterns (e.g. "ללא ציפוי")
    }
    
    # חיפוש קובץ email עם סיומות שונות + fallback לכל קובץ txt בתיקייה
    email_patterns = ['email.txt', 'email', 'Email.txt', 'EMAIL.txt', 'email.text']
    candidate_files = []
    for pattern in email_patterns:
        candidate_files.append(folder_path / pattern)
    # fallback: כל קובץ txt בתיקייה אם לא נמצא אחד מהידועים
    candidate_files.extend(folder_path.glob('*.txt'))

    logger.debug(f"Looking for email files in: {len(candidate_files)} candidate paths...")
    for email_file in candidate_files:
        if not email_file.exists() or not email_file.is_file():
            continue
        
        logger.info(f"✓ Found: {email_file.name}")
        try:
            with open(email_file, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            logger.debug(f"Content length: {len(content)} chars")
                
            # נסה לחלץ subject ו-body
            lines = content.split('\n')
            
            # חיפוש Subject
            subject_idx = None
            for i, line in enumerate(lines):
                line_strip = line.strip()
                if line_strip.lower().startswith('subject:') or line_strip.startswith('נושא'):
                    result['subject'] = line_strip.split(':', 1)[1].strip() if ':' in line_strip else ''
                    subject_idx = i
                    break
            
            # חיפוש From/כתובת שולח
            # תחילה: בדוק אם השורה הראשונה היא כתובת מייל (general case: email address on first line)
            if lines and '@' in lines[0]:
                first_line = lines[0].strip()
                if first_line and '@' in first_line:
                    result['from'] = first_line.replace("כתובת שולח:", "").replace("From:", "").strip()
                    logger.debug(f"[EMAIL] From first line: {result['from']}")
            
            # אם לא נמצא בשורה הראשונה, חיפוש בשורות אחרות עם prefixes
            if not result['from']:
                for i, line in enumerate(lines):
                    line_strip = line.strip()
                    if (line_strip.lower().startswith('from:') or 
                        line_strip.startswith('כתובת שולח:') or
                        line_strip.startswith('שם שולח:')):
                        sender = line_strip.split(':', 1)[1].strip() if ':' in line_strip else ''
                        if sender:
                            result['from'] = sender
                            break
            
            # שאר התוכן הוא הגוף
            if subject_idx is not None and len(lines) > subject_idx + 1:
                result['body'] = '\n'.join(lines[subject_idx + 1:]).strip()
            else:
                result['body'] = content
            
            result['found'] = True
            
            logger.debug(f"Subject: {result['subject'][:50]}...")
            logger.debug(f"Body length: {len(result['body'])} chars")
            
            # חיפוש הנושא האמיתי בגוף המייל (אם המייל הוא replied/forwarded)
            # דוגמה: 'you received an email from tal@chrom.co.il: "FW: הצעת מחיר 5 - 8 פרטים..."'
            
            # חיפוש המרסל האמיתי מהגוף (אם זה forwarded)
            real_from_match = re.search(r'received an email from\s+([^\s:]+)\s*:', result['body'])
            if real_from_match:
                real_from = real_from_match.group(1).strip()
                if real_from and '@' in real_from:  # בדוק שזה כנראה email
                    result['from'] = real_from
                    logger.debug(f"[FOUND] Real sender from body: {real_from}")
            
            # חיפוש הנושא האמיתי בגוף המייל
            real_subject_match = re.search(r'received an email[^:]*:\s*"([^"]+)"', result['body'])
            if not real_subject_match:
                real_subject_match = re.search(r'Subject:\s*([^\n]+)', result['body'], re.IGNORECASE)
            if not real_subject_match:
                real_subject_match = re.search(r'נושא:\s*([^\n]+)', result['body'])
            
            if real_subject_match:
                actual_subject = real_subject_match.group(1).strip()
                logger.debug(f"[FOUND] Actual subject in body: {actual_subject[:60]}...")
                # תמזג את הנושא וגוף - הנושא האמיתי חשוב יותר
                result['subject'] = actual_subject
                result['body'] = result['subject'] + '\n\n' + result['body']
            
            # זיהוי כמויות לשורות פריטים בטקסט (ללא AI)  דומה לטבלה מק"ט/כמות/מחיר
            part_quantities = {}
            
            # חפש דפוס של מספר פריט עם כמויות (כולל מספר כמויות מופרדות בפסיקים)
            # דוגמאות:
            # 85-61-00539-00   25, 40
            # 85-70-00373-00   2,10,20
            for line in lines:
                # דפוס: מספר פריט (עם מקפים/קווים תחתונים) ואחריו כמויות
                # חפש מספר פריט בתחילת השורה או אחרי רווח
                match = re.search(r'\b([A-Za-z0-9][\w\-]{4,})\s+([0-9\s,]+)', line)
                if match:
                    part_raw = match.group(1).strip()
                    qty_text = match.group(2).strip()
                    
                    # חלץ את כל המספרים מטקסט הכמויות (מטפל ב-"25, 40" או "2,10,20")
                    quantities = re.findall(r'\b(\d+)\b', qty_text)
                    
                    if quantities and len(part_raw) >= 4:
                        part_norm = _normalize_item_number(part_raw)
                        if part_norm:
                            # שמור את כל הכמויות
                            if len(quantities) == 1:
                                part_quantities[part_norm] = quantities[0]
                            else:
                                # אם יש מספר כמויות, שמור אותן כרשימה
                                part_quantities[part_norm] = quantities
                            continue
                
                # Fallback: דפוס גמיש יותר - מק"ט ואחריו כמות בודדת
                row_pattern = re.compile(r"\b([A-Za-z0-9][\w\-]{3,})\b[\t\s]+([0-9]{1,5})\b")
                m = row_pattern.search(line)
                if m:
                    part_raw = m.group(1)
                    qty_raw = m.group(2)
                    part_norm = _normalize_item_number(part_raw)
                    if part_norm and qty_raw and part_norm not in part_quantities:
                        part_quantities[part_norm] = qty_raw
                        continue
                        
                # Fallback נוסף: פיצול לפי רווחים/טאבים אם הטוקן הראשון נראה כמו מק"ט ואחריו מספר
                tokens = line.strip().split()
                if len(tokens) >= 2:
                    part_raw = tokens[0]
                    qty_raw = tokens[1].replace(',', '')  # הסר פסיקים
                    if qty_raw.isdigit() and len(part_raw) >= 4:
                        part_norm = _normalize_item_number(part_raw)
                        if part_norm and part_norm not in part_quantities:
                            part_quantities[part_norm] = qty_raw
                            
            result['part_quantities'] = part_quantities
            if part_quantities:
                logger.info(f"Found per-part quantities in email/txt: {len(part_quantities)} rows")
            
            # חילוץ תיאור עבודה כללי מהמייל בהנתן מילות מפתח (ללא AI)
            # מילות מפתח קשורות לציפוי וצביעה (כולל טעויות כתיב נפוצות)
            keywords = [
                'ציפוי', 'צפוי', 'צביעה', 'צבע', 'קשיחים', 'חריצה', 'מילוי', 'מסך', 'קדח',
                'ציפוי_ניקל', 'ציפוי_כרום', 'ציפוי_זהב', 'ציפוי_כסף', 'ציפוי_נחושת',
                'צביעה_ספריי', 'צביעה_מברשת', 'צביעה_טבילה',
                'קרום', 'עמיד', 'מט', 'מבריק', 'חיספוס',
                'בדיקה', 'בדיקת_קשיחות', 'בדיקת_ציפוי', 'דגימה',
                'מהירות_ציפוי', 'עובי_ציפוי', 'תקן', 'ג\'ל',
                # ציפויים ספציפיים
                'פסיבציה', 'passivation', 'אנודייז', 'אנודיזציה', 'anodizing', 'anodize',
                'תמורה', 'כרומי', 'כרומיזציה', 'chromizing', 'chromizing',
                'ציפוי_חשמלי', 'electroplating', 'plating',
                'ציפוי_ניקל_רך', 'ניקל_קשה', 'nickel', 'כרום_קשה', 'chrome', 'hard_chrome',
                'ציפוי_אבץ', 'צפוי_אבץ', 'zinc_plating', 'zinc', 'אבץ',
                'ציפוי_דיכרומט', 'dichromate',
                'ציפוי_סטנד', 'tin_plating', 'tin',
                'התחמצנות', 'oxidation', 'בדיקת_חלודה', 'corrosion_test',
                'בדיקת_ניחור', 'salt_spray', 'rust_test',
                # סימון וחריטה
                'סימון', 'חריטה', 'engraving', 'הדפסה', 'print', 'printing', 'תיוג', 'tag',
                'identification_marking', 'mark', 'marking', 'label', 'מדבקה',
                'epoxy', 'אפוקסי', 'epoxy_paint', 'character_height', 'גובה_תו',
                'line_thickness', 'עובי_קו', 'depth', 'עומק',
                # החדרת קשיחים
                'coil', 'helical', 'insert', 'installation', 'חדרת', 'קשיח', 'כריכה',
                # הסרה וציפוי מחדש
                'הסרה', 'stripping', 'strip', 'מחדש', 'replate', 're-plate', 'replating',
                'removal', 'remove', 'הסר', 'תכליל_ציפוי',
                # טיפולים כימיים וטיפולי חום
                'ניקוי', 'cleaning', 'שטוח', 'polishing', 'cleaning', 'degreasing',
                'גריינדינג', 'grinding', 'חספוס', 'roughness', 'surface_finish',
                'שיחרור_מימן', 'hydrogen', 'degassing', 'heat_treatment', 'בקיאה',
                # הפניות לשרטוט
                'לפי_שרטוט', 'לפי_ציור', 'per_drawing', 'per_specification',
                'שרטוט', 'drawing', 'ציור', 'as_per_drawing'
            ]
            
            email_text_lower = result['body'].lower()
            work_desc_candidates = []
            
            # חפש טקסט שמכיל מילות מפתח
            for keyword in keywords:
                if keyword.replace('_', ' ') in email_text_lower or keyword.replace('_', '') in email_text_lower:
                    # מצא משפט או שורה שמכילה את המילה המפתח
                    sentences = re.split(r'[.!?;\n]', result['body'])
                    for sentence in sentences:
                        sentence_lower = sentence.lower()
                        keyword_clean = keyword.replace('_', ' ').lower()
                        keyword_clean2 = keyword.replace('_', '').lower()
                        
                        # חפש מילה המפתח בעמוד מלוות (ל, עבור, +, גם, כולל וכו')
                        # דוגמאות: "לציפוי", "עבור צביעה", "ציפוי + צביעה", "גם סימון"
                        if (keyword_clean in sentence_lower or keyword_clean2 in sentence_lower or
                            f'ל{keyword_clean}' in sentence_lower or
                            f'עבור {keyword_clean}' in sentence_lower or
                            f'+ {keyword_clean}' in sentence_lower or
                            f'גם {keyword_clean}' in sentence_lower or
                            f'כולל {keyword_clean}' in sentence_lower or
                            (f'כולל' in sentence_lower and (keyword_clean in sentence_lower or keyword_clean2 in sentence_lower))):
                            
                            clean_sentence = sentence.strip()
                            if clean_sentence and len(clean_sentence) > 5 and clean_sentence not in work_desc_candidates:
                                work_desc_candidates.append(clean_sentence)
            
            if work_desc_candidates:
                # תצרף את כל המשפטים שמכילים מילות מפתח
                result['work_description'] = ' + '.join(work_desc_candidates[:5])  # קח עד 5 משפטים
                logger.info(f"Found work description keywords in email: {result['work_description'][:80]}...")
            
            # חילוץ כמויות כלליות באמצעות AI (אם client זמין)
            # חילוץ כמויות כלליות באמצעות AI (אם client זמין)
            if client:
                logger.debug(f"[AI] Extracting general quantities...")
                full_text = result['subject'] + '\n\n' + result['body']
                
                # השתמש ב-AI לחילוץ כמויות מתוך ההקשר
                prompt = load_prompt("07_extract_quantities_from_email")

                try:
                    logger.info(f"[AI] Sending to Azure OpenAI...")
                    email_model, email_max_tokens, email_temperature = _resolve_stage_call_config(
                        STAGE_EMAIL_QUANTITIES,
                        800,
                        0,
                    )

                    content_response = None  # ensure defined for except blocks

                    # o4-mini and other o-series models don't support temperature/seed/response_format
                    api_kwargs = {
                        "model": email_model,
                        "messages": [
                            {"role": "system", "content": prompt},
                            {"role": "user", "content": f"מייל:\n\n{full_text[:2000]}"}  # הגבל ל-2000 תווים
                        ],
                        "max_tokens": email_max_tokens,
                    }

                    # Only add these params for non-reasoning models
                    is_reasoning_model = MODEL_RUNTIME.is_model_reasoning(email_model)
                    if not is_reasoning_model:
                        api_kwargs["temperature"] = email_temperature
                        api_kwargs["seed"] = 12345
                        api_kwargs["response_format"] = RESPONSE_FORMAT

                    response = _chat_create_with_token_compat(client, **api_kwargs)
                    content_response = response.choices[0].message.content or ""
                    content_response = content_response.strip()

                    # o-series reasoning models can return empty content if
                    # reasoning tokens consumed the entire budget.  Retry once
                    # with gpt-4o-mini-email and a simpler prompt.
                    if not content_response and is_reasoning_model:
                        logger.warning(f"⚠️ o4-mini returned empty response — falling back to gpt-4o-mini-email (simple prompt)")
                        fallback_prompt = load_prompt("07b_extract_quantities_fallback")

                        fallback_kwargs = {
                            "model": "gpt-4o-mini-email",
                            "messages": [
                                {"role": "system", "content": fallback_prompt},
                                {"role": "user", "content": f"מייל:\n\n{full_text[:2000]}"},
                            ],
                            "max_tokens": 200,
                            "temperature": 0,
                            "seed": 12345,
                            "response_format": RESPONSE_FORMAT,
                        }
                        response = _chat_create_with_token_compat(client, **fallback_kwargs)
                        content_response = (response.choices[0].message.content or "").strip()
                        # Adapt old flat format → new structured format
                        if content_response:
                            try:
                                # Strip markdown code fences if present
                                fb_text = content_response
                                if fb_text.startswith('```'):
                                    first_nl = fb_text.find('\n')
                                    if first_nl != -1:
                                        fb_text = fb_text[first_nl + 1:]
                                    if fb_text.rstrip().endswith('```'):
                                        fb_text = fb_text.rstrip()[:-3].rstrip()
                                fb_data = json.loads(fb_text)
                                fb_quantities = fb_data.get('quantities', [])
                                if fb_quantities:
                                    content_response = json.dumps({
                                        "part_quantities": {},
                                        "general_quantity": ", ".join(str(q) for q in fb_quantities) if fb_quantities else None,
                                        "quantity_summary": f"כמויות מהמייל: {', '.join(str(q) for q in fb_quantities)}"
                                    }, ensure_ascii=False)
                                    logger.info(f"✓ Fallback found quantities: {fb_quantities}")
                                else:
                                    content_response = ""
                            except json.JSONDecodeError:
                                pass  # will be caught below

                    if not content_response:
                        logger.warning(f"⚠️ AI returned empty response — relying on regex-extracted quantities")
                    else:
                        logger.debug(f"[AI] Response: {content_response[:200]}")
                        # Strip markdown code fences (```json ... ```) if present
                        json_text = content_response
                        if json_text.startswith('```'):
                            # Remove opening fence (```json or ```)
                            first_newline = json_text.find('\n')
                            if first_newline != -1:
                                json_text = json_text[first_newline + 1:]
                            # Remove closing fence
                            if json_text.rstrip().endswith('```'):
                                json_text = json_text.rstrip()[:-3].rstrip()
                        data = json.loads(json_text)

                        # New format: structured quantities
                        part_quantities_ai = data.get('part_quantities', {})
                        general_quantity = data.get('general_quantity', None)
                        quantity_summary = data.get('quantity_summary', '')

                        # Merge AI part_quantities with regex-extracted ones
                        # AI takes priority (more context-aware)
                        # Filter out garbage keys: must contain at least 1 digit
                        # and not be a pure year/date or Hebrew-only word
                        for part_num, qty in part_quantities_ai.items():
                            part_norm = _normalize_item_number(part_num)
                            if not part_norm:
                                continue
                            # Key must contain at least 3 digits to look like a real part number
                            # (I→1/O→0 substitution gives words like "MIL-PRF" → "m1lprf" only 1 digit)
                            if sum(c.isdigit() for c in part_norm) < 3:
                                logger.debug(f"ℹ️ Skipping AI part_quantities key '{part_num}' — <3 digits (not a part number)")
                                continue
                            # Skip pure 4-digit years (2020-2039)
                            if re.fullmatch(r'20[2-3]\d', part_norm):
                                logger.debug(f"ℹ️ Skipping AI part_quantities key '{part_num}' — looks like a year")
                                continue
                            result['part_quantities'][part_norm] = qty

                        # General quantity: can be number ("50"), range ("10-20"),
                        # or comma-separated ("3, 5, 10").
                        # If it contains Hebrew/English words → extract just numbers.
                        if general_quantity:
                            gq_str = str(general_quantity).strip()
                            # Check if it's a "quantity expression" (digits + separators only)
                            # Allows: "50", "10-20", "3, 5, 10", "100/200"
                            gq_qty_only = re.sub(r'[\s,\-/]+', '', gq_str)
                            if gq_qty_only.isdigit():
                                # Pure quantity expression → keep as-is
                                result['general_quantities'] = [gq_str]
                            else:
                                # Text with numbers — extract all numeric values
                                nums = re.findall(r'\d+', gq_str)
                                if nums:
                                    result['general_quantities'] = nums
                                    logger.debug(f"ℹ️ general_quantity was text '{gq_str}' → extracted numbers: {nums}")
                                else:
                                    result['general_quantities'] = []
                                # Move the text to quantity_summary
                                if quantity_summary:
                                    quantity_summary = f"{quantity_summary} | {gq_str}"
                                else:
                                    quantity_summary = gq_str
                        else:
                            # Fallback: if general_quantity is null but quantity_summary has numbers,
                            # extract them so they reach the quantity field
                            if quantity_summary:
                                fallback_nums = re.findall(r'\d+', quantity_summary)
                                # Filter out small numbers that are likely counts ("2 פריטים")
                                qty_nums = [n for n in fallback_nums if int(n) >= 3 or len(fallback_nums) == 1]
                                if qty_nums:
                                    result['general_quantities'] = qty_nums
                                    logger.debug(f"ℹ️ Extracted quantities from summary '{quantity_summary}': {qty_nums}")
                                else:
                                    result['general_quantities'] = []
                            else:
                                result['general_quantities'] = []

                        # Store summary
                        result['quantity_summary'] = quantity_summary

                        # ── Parse work description fields from AI ──
                        part_work_descs = data.get('part_work_descriptions', {})
                        if part_work_descs and isinstance(part_work_descs, dict):
                            for part_num, desc in part_work_descs.items():
                                part_norm = _normalize_item_number(part_num)
                                if part_norm and desc:
                                    result['part_work_descriptions'][part_norm] = str(desc).strip()
                            if result['part_work_descriptions']:
                                logger.info(f"✓ AI found per-part work descriptions: {len(result['part_work_descriptions'])} items")

                        general_work_desc = data.get('general_work_description', None)
                        if general_work_desc:
                            result['general_work_description'] = str(general_work_desc).strip()
                            logger.info(f"✓ AI found general work description: {result['general_work_description'][:80]}")

                        negation = data.get('work_description_negation', None)
                        if negation:
                            result['work_description_negation'] = str(negation).strip()
                            logger.info(f"✓ AI found work description negation: {result['work_description_negation'][:80]}")

                        if result['part_quantities']:
                            logger.info(f"✓ AI found per-part quantities: {len(result['part_quantities'])} items")
                        if result['general_quantities']:
                            logger.info(f"✓ AI found general quantity: {result['general_quantities']}")
                        if quantity_summary:
                            logger.info(f"✓ Summary: {quantity_summary}")
                        if not result['part_quantities'] and not result['general_quantities']:
                            logger.debug(f"ℹ️ No quantities found by AI")
                    
                except json.JSONDecodeError as je:
                    logger.warning(f"✗ JSON parse error: {je}")
                    if content_response is not None:
                        logger.debug(f"Response was: {content_response[:200]}")
                    result['general_quantities'] = []
                except Exception as e:
                    logger.warning(f"✗ Failed to extract quantities with AI: {e}")
                    result['general_quantities'] = []
            else:
                logger.debug(f"[AI DISABLED] No client provided")

            # ── Final sanitization: remove garbage keys from part_quantities ──
            # Keys must look like real part numbers:
            #   - contain at least 3 digits (real part numbers have many digits;
            #     words like "MIL-PRF" → "m1lprf" only get 1 digit from I→1)
            #   - not be a pure 4-digit year (2020-2039)
            # This catches junk from ALL extraction paths (regex + AI).
            raw_pq = result.get('part_quantities', {})
            if raw_pq:
                clean_pq = {}
                for pkey, pval in raw_pq.items():
                    digit_count = sum(c.isdigit() for c in pkey)
                    if digit_count < 3:
                        logger.debug(f"ℹ️ Sanitize: dropping part_quantities key '{pkey}' — only {digit_count} digit(s)")
                        continue
                    if re.fullmatch(r'20[2-3]\d', pkey):
                        logger.debug(f"ℹ️ Sanitize: dropping part_quantities key '{pkey}' — looks like a year")
                        continue
                    clean_pq[pkey] = pval
                dropped = len(raw_pq) - len(clean_pq)
                if dropped:
                    logger.info(f"ℹ️ Sanitized part_quantities: dropped {dropped} garbage key(s), kept {len(clean_pq)}")
                result['part_quantities'] = clean_pq

            break
                
        except Exception as e:
            logger.warning(f"Failed to read email file {email_file}: {e}")
            continue
    
    return result


# ────────────────────────────────────────────────────────────────────────
def _extract_quantities_from_order_pdf(file_path: Path) -> Dict[str, str]:
    """Extract quantities from order PDFs using raw text parsing."""
    results: Dict[str, str] = {}
    
    try:
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return results
            
            # Read up to 5 pages (quotes/POs can span multiple pages)
            full_text = "\n".join(
                (p.extract_text() or "") for p in pdf.pages[:5]
            )
            if not full_text.strip():
                return results
            
            lines = full_text.split('\n')
            part_pattern = r'([0-9]{2}-[0-9]{2}-[0-9]{4,6}-[0-9]{2}|[0-9]{10})(?:\-\d+)?'
            
            for line in lines:
                if not line.strip() or len(line) < 8:
                    continue
                if any(x in line.lower() for x in ['סה"כ', 'total', 'sum', 'page']):
                    continue
                
                for match in re.finditer(part_pattern, line):
                    part_num = match.group(1)
                    norm_part = _normalize_item_number(part_num)
                    if not norm_part or len(norm_part) < 4 or norm_part in results:
                        continue
                    
                    before = line[:match.start()]
                    tokens = re.findall(r'\d+(?:\.\d+)?', before)
                    if not tokens:
                        continue
                    
                    qty = None
                    for tok in reversed(tokens):
                        if '.' in tok and len(tok.split('.')[1]) >= 2:
                            continue
                        try:
                            val = int(float(tok))
                            if 1 <= val <= 10000:
                                qty = str(val)
                                break
                        except (ValueError, TypeError):
                            pass
                    
                    if qty:
                        results[norm_part] = qty
    except Exception as e:
        debug_print(f"[ORDER_PDF_QUANTITY_PARSE] Error: {e}")
        pass
    
    return results


# ────────────────────────────────────────────────────────────────────────
def _extract_text_via_ocr(file_path: Path) -> str:
    """Extract text from PDF using Tesseract OCR when pdfplumber fails."""
    try:
        import pytesseract
        from pdf2image import convert_from_path
        
        logger.info(f"[OCR] Converting {file_path.name} to image...")
        images = convert_from_path(str(file_path), dpi=300, first_page=1, last_page=1)
        
        if not images:
            return ""
        
        # Convert PIL image to numpy array
        image = images[0]
        img_array = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        # Preprocess image for better OCR
        img_gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
        
        # Denoise
        img_denoised = cv2.fastNlMeansDenoising(img_gray, None, h=10, templateWindowSize=7, searchWindowSize=21)
        
        # Enhance contrast with CLAHE
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        img_enhanced = clahe.apply(img_denoised)
        
        # Morphological operations
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        img_enhanced = cv2.morphologyEx(img_enhanced, cv2.MORPH_CLOSE, kernel)
        
        logger.info(f"[OCR] Running Tesseract on preprocessed image...")
        ocr_text = pytesseract.image_to_string(img_enhanced, lang='heb+eng')
        
        return ocr_text.strip() if ocr_text else ""
    
    except Exception as e:
        logger.warning(f"[OCR] Failed: {e}")
        return ""


# ────────────────────────────────────────────────────────────────────────
_ITEM_COL_PATTERNS = [
    r'מק"?ט', r'מקט', r'פריט', r'מספר\s*חלק', r'מספר\s*פריט',
    r'part[\s_-]*n', r'p[./]?n', r'item[\s_-]*n', r'drawing',
    r'catalog', r'קטלוג', r'חלק',
]
_QTY_COL_PATTERNS = [
    r'כמות', r'כמ\'', r'qty', r'quantity', r'quant', r'count', r'יחידות',
]
_DESC_COL_PATTERNS = [
    r'תיאור', r'desc', r'פירוט', r'שם\s*פריט', r'item\s*name',
]


def _find_column(columns: List[str], patterns: list) -> Optional[str]:
    """מצא עמודה לפי רשימת תבניות regex — מחזיר את שם העמודה הראשון שמתאים."""
    for col in columns:
        col_lower = str(col).strip().lower()
        for pat in patterns:
            if re.search(pat, col_lower):
                return col
    return None


def _extract_item_details_from_excel(file_path: Path) -> Dict[str, Any]:
    """
    חילוץ מספרי פריט + כמויות (+ תיאור אופציונלי) מקובץ Excel.
    מחזיר dict באותו פורמט של item_details הקיים:
        { item_number: { 'quantities': [...], 'work_description': str, 'original_item_number': str } }
    """
    details: Dict[str, Any] = {}
    try:
        # קרא את כל הגיליונות
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name)
            except Exception:
                continue

            if df.empty or len(df.columns) < 2:
                continue

            # --- אם אין כותרות שמתאימות, נסה לזהות שורת כותרת בתוך הנתונים ---
            item_col = _find_column(df.columns.tolist(), _ITEM_COL_PATTERNS)
            qty_col = _find_column(df.columns.tolist(), _QTY_COL_PATTERNS)

            if not item_col or not qty_col:
                # חפש בשורות הראשונות (עד 5) שורה שמכילה כותרות
                for row_idx in range(min(5, len(df))):
                    row_vals = [str(v).strip() for v in df.iloc[row_idx]]
                    test_item = _find_column(row_vals, _ITEM_COL_PATTERNS)
                    test_qty = _find_column(row_vals, _QTY_COL_PATTERNS)
                    if test_item and test_qty:
                        # מצאנו שורת כותרות — בנה מחדש את ה-DF
                        df.columns = row_vals
                        df = df.iloc[row_idx + 1:].reset_index(drop=True)
                        item_col = test_item
                        qty_col = test_qty
                        break

            if not item_col or not qty_col:
                continue  # לא מצאנו עמודות מתאימות בגיליון הזה

            desc_col = _find_column(df.columns.tolist(), _DESC_COL_PATTERNS)

            for _, row in df.iterrows():
                raw_item = str(row.get(item_col, '')).strip()
                if not raw_item or raw_item.lower() in ('nan', '', 'none'):
                    continue

                item_num = _normalize_item_number(raw_item)
                if not item_num:
                    continue

                raw_qty = str(row.get(qty_col, '')).strip()
                # נקה ערכי NaN / ריקים
                if raw_qty.lower() in ('nan', '', 'none'):
                    raw_qty = ''

                work_desc = ''
                if desc_col is not None:
                    raw_desc = str(row.get(desc_col, '')).strip()
                    if raw_desc.lower() not in ('nan', '', 'none'):
                        work_desc = raw_desc

                if item_num not in details:
                    details[item_num] = {
                        'quantities': [],
                        'work_description': '',
                        'original_item_number': raw_item,
                    }

                if raw_qty and raw_qty not in details[item_num]['quantities']:
                    details[item_num]['quantities'].append(raw_qty)

                if work_desc and len(work_desc) > len(details[item_num]['work_description']):
                    details[item_num]['work_description'] = work_desc

        if details:
            logger.info(f"Excel extraction from {file_path.name}: {len(details)} items found")

    except Exception as e:
        logger.warning(f"Failed to extract from Excel {file_path.name}: {e}")

    return details


# ────────────────────────────────────────────────────────────────────────
def _extract_item_details_from_documents(folder_path: Path, file_classifications: List[Dict], 
                                         client: AzureOpenAI, email_data: Optional[Dict] = None) -> Dict[str, Any]:
    """
    חילוץ כמויות ותיאורי עבודה מהזמנות/הצעות מחיר בלבד (ללא PL)
    אם לא מצאה תיאור בטבלה, תשתמש בתיאור כללי מהמייל (אם קיים)
    
    Args:
        folder_path: נתיב לתקייה
        file_classifications: רשימת קבצים מסווגים
        client: Azure OpenAI client
        email_data: דיקשונרי עם נתוני מייל (אופציונלי, לשימוש כ-fallback)
        
    Returns:
        item_details: Dict[item_number, {'quantities': [...], 'work_description': str}]
    """
    item_details = {}

    # Define the analysis prompt once to avoid scope issues
    prompt = load_prompt("08_extract_item_details_from_orders")
    
    # מצא קבצי הזמנה, הצעת מחיר והצעות בלבד (לא PL!)
    relevant_files = [
        fc for fc in file_classifications 
        if fc.get('file_type') in ['PURCHASE_ORDER', 'QUOTE', 'INVOICE']
    ]
    
    for doc_idx, fc in enumerate(relevant_files, 1):
        file_path = fc.get('file_path')
        if not file_path:
            continue
            
        doc_type = fc.get('file_type')
        
        try:
            if file_path.suffix.lower() == '.pdf':
                with pdfplumber.open(file_path) as pdf:
                    if len(pdf.pages) == 0:
                        continue
                    # Extract image from first page with ULTRA high resolution
                    page = pdf.pages[0]
                    img = page.to_image(resolution=200)  # מאוזן — מספיק לקריאת טבלאות
                    img_buffer = io.BytesIO()
                    img.original.save(img_buffer, format='PNG')
                    img_bytes = img_buffer.getvalue()
            else:
                # Image file
                with open(file_path, 'rb') as f:
                    img_bytes = f.read()
            
            #  חתוך את אזור הטבלה הרלוונטי
            logger.info(f"Enhanced processing: cropping table area for {file_path.name}...")
            try:
                nparr = np.frombuffer(img_bytes, np.uint8)
                img_full = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                
                if img_full is not None:
                    full_height, full_width = img_full.shape[:2]
                    
                    # חתוך אזור גדול שמכיל את כל הטבלה (90% רוחב x 85% גובה)
                    # מתחיל מ-10% מלמעלה כדי לכלול כותרות
                    # לוקח עד כמעט הסוף כדי לכלול את השורה התחתונה עם הכמויות
                    start_y = int(full_height * 0.10)  # התחל 10% מלמעלה
                    start_x = int(full_width * 0.05)   # התחל 5% משמאל
                    table_height = int(full_height * 0.85)  # קח 85% מהגובה
                    table_width = int(full_width * 0.90)    # קח 90% מהרוחב
                    
                    table_region = img_full[start_y:start_y+table_height, start_x:start_x+table_width]
                    
                    # הגדל פי 2 — מספיק לקריאת טבלאות
                    img_enlarged = cv2.resize(table_region, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
                    
                    #  שיפורי עיבוד תמונה מתקדמים
                    # 1. המרה לגווני אפור
                    gray = cv2.cvtColor(img_enlarged, cv2.COLOR_BGR2GRAY)
                    
                    # 2. Denoising חזק יותר - הסרת רעש
                    denoised = cv2.fastNlMeansDenoising(gray, None, h=12, templateWindowSize=7, searchWindowSize=21)
                    
                    # 3. Double sharpening - חידוד כפול לבהירות מקסימלית
                    kernel_sharp = np.array([[-1, -1, -1],
                                            [-1,  9, -1],
                                            [-1, -1, -1]])
                    sharpened = cv2.filter2D(denoised, -1, kernel_sharp)
                    sharpened2 = cv2.filter2D(sharpened, -1, kernel_sharp)  # חידוד נוסף
                    
                    # 4. Contrast enhancement מוגבר - שיפור ניגודיות
                    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))  # clipLimit גבוה יותר
                    enhanced = clahe.apply(sharpened2)
                    
                    # 5. Morphological operations - ניקוי רעשים קטנים
                    kernel_morph = np.ones((2, 2), np.uint8)
                    morphed = cv2.morphologyEx(enhanced, cv2.MORPH_CLOSE, kernel_morph)
                    
                    # 6. Adaptive Thresholding משופר - סף אדפטיבי
                    binary = cv2.adaptiveThreshold(morphed, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                                   cv2.THRESH_BINARY, 13, 3)  # פרמטרים משופרים
                    
                    # המר חזרה ל-BGR
                    final_image = cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)
                    
                    # Encode, 5: True
                    _, buffer_final = cv2.imencode('.png', final_image)
                    img_bytes = buffer_final.tobytes()
                    
                    logger.info(f"Enhanced TABLE: 2x zoom + processing ({img_enlarged.shape[1]}x{img_enlarged.shape[0]}px)")
                else:
                    logger.debug(f"Could not decode image, using original")
                    
            except Exception as crop_err:
                logger.debug(f"Enhancement failed, using original: {crop_err}")
            
            if not img_bytes:
                continue
            
            # Encode to base64
            image_b64 = base64.b64encode(img_bytes).decode('utf-8')
            
            # בקש מ-GPT לחלץ פרטי פריטים (מסמכי הזמנה/הצעת מחיר)
            # prompt מוגדר למעלה כדי להבטיח שהוא תמיד זמין
            messages = [
                {"role": "system", "content": prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "חלץ את פרטי הפריטים, הכמויות והתיאורים מהמסמך:"},
                        {"type": "image_url", "image_url": {
                            "url": f"data:image/png;base64,{image_b64}",
                            "detail": "high"
                        }}
                    ]
                }
            ]

            # Call Vision API with automatic content filter retry
            logger.info(f"Sending PO/Quote to Vision API ({len(img_bytes)/1024:.0f}KB image)...")
            response = _call_vision_api_with_retry(client, messages, max_tokens=1500, temperature=0, stage_num=STAGE_ORDER_ITEM_DETAILS)
            if response is None:
                logger.warning(f"Vision API failed for {file_path.name}")
                continue
            logger.info(f"Vision API response received for {file_path.name}")
            
            content = response.choices[0].message.content
            
            # Parse JSON
            try:
                # Try to extract JSON block if response contains extra text
                json_match = re.search(r'\{[\s\S]*\}', content)
                if json_match:
                    data = json.loads(json_match.group())
                else:
                    data = json.loads(content)
                
                items = data.get('items', [])

                # Fallback: try deterministic table qty extraction from the PDF
                table_qty_map = {}
                if file_path.suffix.lower() == '.pdf':
                    table_qty_map = _extract_quantities_from_order_pdf(file_path)
                
                for item_idx, item in enumerate(items, 1):
                    item_num = str(item.get('item_number', '')).strip()
                    quantities = item.get('quantities', [])
                    work_desc = (item.get('work_description') or '').strip()
                    
                    if not item_num:
                        continue
                    
                    # נרמל את מספר הפריט (הסר [D] וכו')
                    item_num_normalized = _normalize_item_number(item_num)
                    
                    # Override quantities from table fallback if available
                    fallback_qty = table_qty_map.get(item_num_normalized)
                    if fallback_qty and (not quantities or quantities[0] in ['1', '2', '3', '4', '5']):
                        quantities = [fallback_qty]

                    # אחד מידע לפי מספר פריט מנורמל
                    if item_num_normalized not in item_details:
                        item_details[item_num_normalized] = {
                            'quantities': [],
                            'work_description': '',
                            'original_item_number': item_num
                        }
                    
                    # הוסף כמויות (ללא כפילויות)
                    for qty in quantities:
                        qty_str = str(qty).strip()
                        if qty_str and qty_str not in item_details[item_num_normalized]['quantities']:
                            item_details[item_num_normalized]['quantities'].append(qty_str)
                    
                    # עדכן תיאור (קח את המפורט ביותר)
                    if work_desc:
                        if not item_details[item_num_normalized]['work_description']:
                            item_details[item_num_normalized]['work_description'] = work_desc
                        elif len(work_desc) > len(item_details[item_num_normalized]['work_description']):
                            item_details[item_num_normalized]['work_description'] = work_desc
                
                logger.info(f"Extracted {len(items)} items from document")
                
            except json.JSONDecodeError as e:
                logger.warning(f"Failed to parse JSON: {e}")
                logger.debug(f"Response length: {len(content)} chars")
                logger.debug(f"First 500 chars: {content[:500]}")
                continue
                
        except Exception as e:
            logger.warning(f"Failed to process {file_path.name}: {e}")
            continue
    
    if relevant_files:
        logger.info(f"Items from PDF/image documents: {len(item_details)}")
    else:
        logger.info("No PDF/image order/quote documents found — checking Excel files...")

    # ── חילוץ כמויות מקבצי Excel מצורפים ──
    excel_files = [
        fc for fc in file_classifications
        if Path(fc.get('file_path', '')).suffix.lower() in {'.xls', '.xlsx'}
    ]
    for fc in excel_files:
        fp = Path(fc.get('file_path', ''))
        if not fp.exists():
            continue
        excel_details = _extract_item_details_from_excel(fp)
        for item_num, info in excel_details.items():
            if item_num not in item_details:
                item_details[item_num] = info
            else:
                # מיזוג כמויות (ללא כפילויות)
                for qty in info.get('quantities', []):
                    if qty not in item_details[item_num]['quantities']:
                        item_details[item_num]['quantities'].append(qty)
                # תיאור — קח את הארוך יותר
                if len(info.get('work_description', '')) > len(item_details[item_num].get('work_description', '')):
                    item_details[item_num]['work_description'] = info['work_description']

    if excel_files:
        logger.info(f"Total unique items after Excel merge: {len(item_details)}")

    # Fallback: אם לא מצאנו תיאור בטבלה, נסה להשתמש בתיאור כללי מהמייל
    if email_data and email_data.get('work_description'):
        email_work_desc = email_data.get('work_description', '')
        items_without_desc = [item for item, details in item_details.items() 
                              if not details['work_description']]
        
        if items_without_desc:
            logger.info(f"Using email work_description for {len(items_without_desc)} items without table description")
            for item in items_without_desc:
                item_details[item]['work_description'] = email_work_desc
    
    for item_num, details in item_details.items():
        logger.debug(f"{item_num}: Qty={details['quantities']} | Desc='{details['work_description'][:40]}...'")
    
    return item_details
