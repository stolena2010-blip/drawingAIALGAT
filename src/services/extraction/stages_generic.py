"""
Generic Drawing Extraction Stages (non-customer-specific)
==========================================================
Stages: identify_drawing_layout, extract_basic_info, extract_processes_info,
        validate_notes_before_stage5, extract_notes_text, calculate_geometric_area

Extracted from customer_extractor_v3_dual.py
"""
import json
import os
import re
from typing import Optional, Tuple, Dict
from openai import AzureOpenAI

from src.utils.logger import get_logger
from src.utils.prompt_loader import load_prompt
from src.core.constants import (
    STAGE_LAYOUT, STAGE_BASIC_INFO, STAGE_PROCESSES, STAGE_NOTES, STAGE_AREA,
)
from src.services.ai.vision_api import (
    _resolve_stage_call_config,
    _chat_create_with_token_compat,
    _call_vision_api_with_retry,
)
from src.services.extraction.ocr_engine import debug_print

logger = get_logger(__name__)


# 

def identify_drawing_layout(image_b64: str, client: AzureOpenAI) -> Tuple[Optional[str], int, int]:
    """
    זיהוי מבנה ופריסת השרטוט - איפה נמצאים השדות השונים
    
    Returns:
        (layout_description, input_tokens, output_tokens)
    """
    prompt = load_prompt("01_identify_drawing_layout")

    messages = [
        {"role": "system", "content": prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": "נתח את מבנה השרטוט - איפה נמצאים השדות:"},
                {"type": "image_url", "image_url": {
                    "url": f"data:image/png;base64,{image_b64}",
                    "detail": "high"
                }}
            ]
        }
    ]
    
    # Call Vision API with automatic content filter retry
    response = _call_vision_api_with_retry(client, messages, max_tokens=250, temperature=0, stage_num=STAGE_LAYOUT)
    if response is None:
        logger.error(f"Layout detection failed")
        return "unknown", 0, 0
    
    content = response.choices[0].message.content
    usage = response.usage
    
    try:
        data = json.loads(content)
    except Exception as e:
        logger.debug(f"Handled: {e}")
        json_match = re.search(r'\{.*\}', content, re.DOTALL)
        if json_match:
            data = json.loads(json_match.group())
        else:
            return "unknown", usage.prompt_tokens, usage.completion_tokens
        
    # Build compact layout description
    layout_desc = f"TB:{data.get('title_block_location', 'N/A')}-{data.get('title_block_size', 'N/A')} | " \
                 f"PN:{data.get('part_number_location', 'N/A')} | " \
                 f"DN:{data.get('drawing_number_location', 'N/A')} | " \
                 f"MAT:{data.get('material_location', 'N/A')} | " \
                 f"NOTES:{data.get('notes_location', 'N/A')} | " \
                 f"Pattern:{data.get('layout_pattern', 'N/A')}"
        
    return layout_desc, usage.prompt_tokens, usage.completion_tokens


# 
#  חילוץ מידע בשני שלבים
# 

def extract_basic_info(text_snippet: str, image_b64: str, client: AzureOpenAI) -> Tuple[Optional[Dict], int, int]:
    """
    שלב 1: חילוץ מידע בסיסי - קצר וממוקד
    """
    prompt = load_prompt("02_extract_basic_info")
    
    messages = [
        {"role": "system", "content": prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": (
                    f" אנא קרא את התמונה בעיון ו**הסתמך על מה שאתה רואה בתמונה** - לא על ה-OCR!\n\n"
                    f" ה-OCR לפעמים לא מדויק - אם הטקסט למטה נראה מעורבב או לא קריא,\n"
                    f"התעלם ממנו וקרא ישירות מהתמונה!\n\n"
                    f" OCR Text (ייתכן שלא מדויק):\n"
                    f"{'='*60}\n"
                    f"{text_snippet[:1500] if text_snippet and text_snippet.strip() else 'OCR failed - read from image only'}\n"
                    f"{'='*60}\n\n"
                    f" **חשוב במיוחד**:\n"
                    f"1. הסתכל בתמונה ב-TITLE BLOCK (בדרך כלל בפינה הימנית התחתונה)\n"
                    f"2. חפש את השדה \"PART NO.\", \"PART NUMBER:\", או \"CATALOG NO.\"\n"
                    f"3. המספר שמופיע מיד אחרי - זה מספר הפריט!\n"
                    f"4. חפש בנפרד את \"DRAWING NO.\" / \"DWG NO.\" - זה מספר השרטוט (שדה אחר!)\n"
                    f"5. העתק את כל מספר בדיוק כפי שרואה בתמונה\n"
                    f"6. גם מספר של 10 ספרות כמו \"0948326040\" - זה מספר פריט תקין!\n\n"
                    f"חלץ מידע מה-TITLE BLOCK בתמונה."
                )},
                {"type": "image_url", "image_url": {
                    "url": f"data:image/png;base64,{image_b64}",
                    "detail": "high"  # Request high-res image analysis
                }}
            ]
        }
    ]
    
    # Call Vision API with automatic content filter retry
    response = _call_vision_api_with_retry(client, messages, max_tokens=250, temperature=0, stage_num=STAGE_BASIC_INFO)
    if response is None:
        return None, 0, 0
    
    content = response.choices[0].message.content
    usage = response.usage
    
    # Parse JSON
    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        json_match = re.search(r'\{.*\}', content, re.DOTALL)
        if json_match:
            json_str = json_match.group()
            json_str = re.sub(r',(\s*[}\]])', r'\1', json_str)
            data = json.loads(json_str)
        else:
            logger.error(f"Failed to parse JSON from Vision API")
            return None, 0, 0
    
    #  POST-PROCESSING: Strip spaces from all ID fields immediately
    for _field in ('part_number', 'drawing_number', 'revision'):
        if data.get(_field) and isinstance(data[_field], str):
            data[_field] = data[_field].strip().replace(' ', '')

    # Reject DATE values mistakenly extracted as part/drawing numbers
    # Patterns: DD.MM.YY, DD.MM.YYYY, DD/MM/YY, DD-MM-YY, YYYY-MM-DD etc.
    _DATE_RE = re.compile(
        r'^(?:'
        r'\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}'
        r'|\d{4}[./\-]\d{1,2}[./\-]\d{1,2}'
        r')$'
    )
    for _field in ('part_number', 'drawing_number'):
        val = data.get(_field)
        if val and isinstance(val, str) and _DATE_RE.match(val.strip()):
            logger.info(f"{_field} '{val}' looks like a DATE — removed")
            data[_field] = None

    # Reject all-zeros placeholder values (e.g. CATALOG NO. = "000000000")
    for _field in ('part_number', 'drawing_number'):
        val = data.get(_field)
        if val and isinstance(val, str) and re.match(r'^0+$', val.strip()):
            logger.info(f"{_field} '{val}' is all zeros (placeholder) — removed")
            data[_field] = None

    #  POST-PROCESSING VALIDATIONS
    validation_warnings = []
    
    # 1. Validate part_number - check if it contains full words
    if data.get('part_number'):
        part_num = str(data['part_number']).strip()
        part_upper = part_num.upper()
        
        # NOTE: CAGE CODE detection moved to customer_extractor_v3_dual.py
        # (FINAL VALIDATION section) where filename is available for validation.
        # This prevents false positives like AL002, AB123 etc. that ARE in filename.
        
        # Only remove if part_number literally contains the word "CAGE CODE"
        if 'CAGE' in part_upper and 'CAGE CODE' in part_upper:
            validation_warnings.append("Part# contains 'CAGE CODE' text")
            data['part_number'] = None
            logger.info(f"Part number contains 'CAGE CODE' text - removed as invalid")
        
        # Check for common description words that shouldn't be in part numbers
        elif data.get('part_number'):  # Only if not already nulled by CAGE check
            description_words = ['BRACKET', 'HOUSING', 'COVER', 'PLATE', 'ASSEMBLY', 'MOUNT', 'SUPPORT', 
                               'ADAPTER', 'CONNECTOR', 'FRAME', 'PANEL', 'BODY', 'CAP', 'RING', 'FLANGE',
                               'SHAFT', 'PIN', 'SCREW', 'BOLT', 'WASHER', 'SPRING', 'GASKET', 'SEAL']
            
            for word in description_words:
                if word in part_upper:
                    # Probably swapped - move to item_name
                    if not data.get('item_name') or data.get('item_name') == 'null':
                        data['item_name'] = part_num
                    data['part_number'] = None
                    validation_warnings.append(f"Part# '{part_num}' is description")
                    logger.info(f"Part number '{part_num}' looks like a description - moved to item_name")
                    break
    
    # 2. Validate drawing_number - check if it's CAGE CODE
    if data.get('drawing_number'):
        dwg_num = str(data['drawing_number']).strip()
        dwg_upper = dwg_num.upper()
        
        is_cage_code_dwg = False
        
        # Check: Exactly 5 alphanumeric characters with both letters AND digits
        if len(dwg_num) == 5 and dwg_num.isalnum():
            has_letter = bool(re.search(r'[A-Z]', dwg_upper))
            has_digit = bool(re.search(r'[0-9]', dwg_num))
            
            if has_letter and has_digit:
                is_cage_code_dwg = True
                validation_warnings.append(f"Drawing# '{dwg_num}' is CAGE CODE format")
                data['drawing_number'] = None
                logger.info(f"Drawing# '{dwg_num}' detected as CAGE CODE (5-char alphanumeric) - removed!")
        
        # Also check if contains "CAGE" keyword
        if not is_cage_code_dwg and 'CAGE' in dwg_upper:
            validation_warnings.append("Drawing# contains CAGE CODE keyword")
            data['drawing_number'] = None
            logger.info(f"Drawing number contains 'CAGE' - removed as invalid")
    
    # 3. Validate customer_name - should not be CAGE CODE
    if data.get('customer_name'):
        cust_name = str(data['customer_name'])
        if 'CAGE' in cust_name.upper() or re.match(r'^\d{5}$', cust_name.strip()):
            validation_warnings.append("Customer name looks like CAGE CODE")
            data['customer_name'] = None
            logger.info(f"Customer name is CAGE CODE - removed")
    
    # Add all validation warnings to the data (will be updated with confidence later)
    if validation_warnings:
        data['validation_warnings'] = '; '.join(validation_warnings)
    else:
        data['validation_warnings'] = ""
    
    # Mark that we did NOT explicitly search for P.N. field (for non-RAFAEL drawings)
    # This allows fallback to drawing_number if part_number is missing
    data['_searched_for_pn_field'] = False
    
    return data, usage.prompt_tokens, usage.completion_tokens


def extract_processes_info(text_snippet: str, image_b64: str, client: AzureOpenAI,
                           additional_images: list = None) -> Tuple[Optional[Dict], int, int]:
    """
    שלב 2: חילוץ תהליכים ומפרטים - מפורט יותר
    """
    prompt = load_prompt("03_extract_processes")
    
    content = [
        {"type": "text", "text": f"TEXT:\n{text_snippet}\n\nחלץ תהליכים ומפרטים."},
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}}
    ]
    # Add additional pages (multi-page drawings)
    if additional_images:
        for pg_b64 in additional_images:
            content.append(
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{pg_b64}"}}
            )

    messages = [
        {"role": "system", "content": prompt},
        {"role": "user", "content": content}
    ]
    
    # Call Vision API with automatic content filter retry
    response = _call_vision_api_with_retry(client, messages, max_tokens=400, temperature=0, stage_num=STAGE_PROCESSES)
    
    if response is None:
        return None, 0, 0
    
    content = response.choices[0].message.content
    usage = response.usage
    
    # Parse JSON
    try:
        data = json.loads(content)
        return data, usage.prompt_tokens, usage.completion_tokens
    except json.JSONDecodeError as e:
        logger.debug(f"Handled: {e}")
        json_match = re.search(r'\{.*\}', content, re.DOTALL)
        if json_match:
            json_str = json_match.group()
            json_str = re.sub(r',(\s*[}\]])', r'\1', json_str)
            try:
                data = json.loads(json_str)
                return data, usage.prompt_tokens, usage.completion_tokens
            except Exception as e2:
                return None, 0, 0
        else:
            return None, 0, 0


def validate_notes_before_stage5(notes_text: Optional[str], debug_prefix: str = "") -> Tuple[Optional[str], dict]:
    """
    תיקוף מפורט של טקסט NOTES לפני שליחה לשלב 5
    
    בודק: encoding, תווים לא תקינים, תוכן משמעותי, אורך תקבול
    """
    report = {
        "is_valid": False,
        "issues": [],
        "char_count": 0,
        "line_count": 0,
        "has_hebrew": False,
        "has_numbers": False,
        "encoding_ok": True,
        "warnings": []
    }
    
    # בדיקה 1: האם יש טקסט בכלל
    if not notes_text or not isinstance(notes_text, str):
        report["issues"].append(f"Invalid type: {type(notes_text).__name__}")
        return None, report
    
    notes_text = notes_text.strip()
    if len(notes_text) == 0:
        report["issues"].append("Empty after strip")
        return None, report
    
    # בדיקה 2: אורך בסיסי
    report["char_count"] = len(notes_text)
    report["line_count"] = len(notes_text.split('\n'))
    
    if len(notes_text) < 20:
        report["warnings"].append(f"Very short ({len(notes_text)} chars)")
    if len(notes_text) > 5000:
        report["warnings"].append(f"Very long ({len(notes_text)} chars)")
    
    # בדיקה 3: Encoding
    try:
        notes_text.encode('utf-8')
        report["encoding_ok"] = True
    except Exception as e:
        logger.debug(f"Handled: {e}")
        report["issues"].append(f"Encoding error: {str(e)[:50]}")
        report["encoding_ok"] = False
        return None, report
    
    # בדיקה 4: בדוק תווים בעיתיים
    problematic_chars = []
    for i, char in enumerate(notes_text):
        code = ord(char)
        if code < 32 and char not in '\n\t\r':
            problematic_chars.append(f"ctrl-{code}@{i}")
    
    if problematic_chars:
        report["warnings"].append(f"Found {len(problematic_chars)} problematic chars")
    
    # בדיקה 5: תוכן
    has_hebrew = any('\u0590' <= ch <= '\u05FF' for ch in notes_text)
    has_english = any(ch.isalpha() and ord(ch) < 128 for ch in notes_text)
    has_numbers = any(ch.isdigit() for ch in notes_text)
    
    report["has_hebrew"] = has_hebrew
    report["has_numbers"] = has_numbers
    
    # סיכום
    report["is_valid"] = len(report["issues"]) == 0
    
    return (notes_text if report["is_valid"] else None), report


def extract_notes_text(text_snippet: str, image_b64: str, client: AzureOpenAI,
                       additional_images: list = None) -> Tuple[Optional[str], int, int]:
    """
    שלב 3: חילוץ טקסט NOTES מלא בלבד
    """
    prompt = load_prompt("04_extract_notes_text")
    
    content = [
        {"type": "text", "text": f"OCR:\n{text_snippet}\n\nחלץ NOTES מהשרטוט."},
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}", "detail": "high"}}
    ]
    # Add additional pages (multi-page drawings)
    if additional_images:
        for pg_b64 in additional_images:
            content.append(
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{pg_b64}"}}
            )

    messages = [
        {"role": "system", "content": prompt},
        {"role": "user", "content": content}
    ]
    
    # Call Vision API with automatic content filter retry
    response = _call_vision_api_with_retry(client, messages, max_tokens=800, temperature=0, stage_num=STAGE_NOTES)
    if response is None:
        return None, 0, 0
    
    usage = response.usage
    json_str = response.choices[0].message.content.strip()
    json_str = re.sub(r',(\s*[}\]])', r'\1', json_str)
    
    try:
        data = json.loads(json_str)
        notes = data.get('notes_full_text')
        return notes, usage.prompt_tokens, usage.completion_tokens
    except json.JSONDecodeError as e:
        logger.debug(f"Handled: {e}")
        match = re.search(r'\{.*\}', json_str, re.DOTALL)
        if match:
            try:
                data = json.loads(match.group())
                notes = data.get('notes_full_text')
                return notes, usage.prompt_tokens, usage.completion_tokens
            except Exception as e2:
                return None, usage.prompt_tokens, usage.completion_tokens
        else:
            return None, usage.prompt_tokens, usage.completion_tokens


def calculate_geometric_area(text_snippet: str, image_b64: str, client: AzureOpenAI) -> Tuple[Optional[str], int, int]:
    """
    שלב 4: הערכה גסה של שטח החלק
    """
    prompt = load_prompt("05_calculate_geometric_area")
    
    messages = [
        {"role": "system", "content": prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": f"TEXT:\n{text_snippet}\n\nזהה צורה וחשב שטח."},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}}
            ]
        }
    ]
    
    # Call Vision API with automatic content filter retry
    response = _call_vision_api_with_retry(client, messages, max_tokens=300, temperature=0, stage_num=STAGE_AREA)
    if response is None:
        return None, 0, 0
    
    usage = response.usage
    json_str = response.choices[0].message.content.strip()
    json_str = re.sub(r',(\s*[}\]])', r'\1', json_str)
    
    try:
        data = json.loads(json_str)
        # Combine the geometric analysis into a readable format
        if data.get('geometric_area'):
            area_info = data['geometric_area']
            # Keep the old field name for compatibility
            return area_info, usage.prompt_tokens, usage.completion_tokens
        else:
            return None, usage.prompt_tokens, usage.completion_tokens
    except Exception as e:
        logger.debug(f"Handled: {e}")
        match = re.search(r'\{.*\}', json_str, re.DOTALL)
        if match:
            data = json.loads(match.group())
            if data.get('geometric_area'):
                return data['geometric_area'], usage.prompt_tokens, usage.completion_tokens
        return None, usage.prompt_tokens, usage.completion_tokens
