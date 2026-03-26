"""
RAFAEL-specific Drawing Extraction Stages
===========================================
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
    STAGE_BASIC_INFO, STAGE_VALIDATION, RESPONSE_FORMAT,
)
from src.services.ai.vision_api import (
    _resolve_stage_call_config,
    _chat_create_with_token_compat,
    _call_vision_api_with_retry,
)
from src.services.extraction.ocr_engine import debug_print, MultiOCREngine
from src.services.extraction.stages_generic import (
    identify_drawing_layout,
    extract_basic_info,
    extract_processes_info,
    extract_notes_text,
    calculate_geometric_area,
)

logger = get_logger(__name__)


# בעתיד ניתן לשנות את הפרומפטים רק כאן ללא השפעה על שאר הלקוחות

def identify_drawing_layout_rafael(image_b64: str, client: AzureOpenAI) -> Tuple[Optional[str], int, int]:
    """
    זיהוי מבנה שרטוט - גרסת רפאל
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי להתאים לפורמט ספציפי של רפאל
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return identify_drawing_layout(image_b64, client)


def extract_basic_info_rafael(text_snippet: str, image_b64: str, client: AzureOpenAI, ocr_engine: Optional[MultiOCREngine] = None, pdfplumber_text: str = '', email_parts_hint: str = '', filename: str = '') -> Tuple[Optional[Dict], int, int]:
    """
    שלב 1 לרפאל: חילוץ מידע בסיסי - מותאם לפורמט RAFAEL
    משתמש ב-Title Block מוגדל לדיוק מקסימלי
    """
    
    #  Create enlarged Title Block version if OCR engine available
    image_to_use = image_b64
    tb_location = "original"
    
    if ocr_engine:
        try:
            # Decode base64 to bytes
            import base64
            img_bytes = base64.b64decode(image_b64)
            
            # Extract and enlarge title block (3x for RAFAEL)
            tb_bytes, tb_location = ocr_engine.extract_title_block_region(img_bytes, enlarge_factor=3.5)
            
            # Re-encode to base64
            image_to_use = base64.b64encode(tb_bytes).decode('utf-8')
            
        except Exception as e:
            logger.error(f"Title Block extraction failed, using full image: {e}")
    
    prompt = load_prompt("06_extract_basic_info_rafael")

    # ─── Extract relevant Title Block lines from pdfplumber text ───
    pdfplumber_hint = ''
    pdf_source = pdfplumber_text or text_snippet or ''
    if pdf_source and pdf_source.strip():
        def _dedup_pdf(line):
            s = line.strip()
            if len(s) >= 6 and len(s) % 2 == 0:
                if all(s[i] == s[i+1] for i in range(0, len(s)-1, 2)):
                    return s[::2]
            return s

        tb_keywords = ['P.N', 'PN ', 'PART NO', 'DRAWING NO', 'DWG NO',
                       'REV ', 'TITLE', 'MMA', 'SCALE', 'SIZE']
        relevant = []
        for line in pdf_source.split('\n'):
            fixed = _dedup_pdf(line)
            upper = fixed.upper()
            # Include lines with TB keywords
            if any(kw in upper for kw in tb_keywords):
                relevant.append(fixed.strip())
            # Also include standalone numeric identifiers (4+ digits,
            # likely drawing numbers for RAFAEL variant without labels)
            elif re.match(r'^\s*\d{4,}\s*$', fixed):
                relevant.append(fixed.strip())
        pdfplumber_hint = '\n'.join(relevant[:10])

    # ─── Build filename hint text ───
    filename_hint_text = ''
    if filename:
        filename_hint_text = (
            f"📁 **שם הקובץ**: {filename}\n"
            f"   ⚠️ שם הקובץ לעתים קרובות מכיל את מספר השרטוט — השתמש כ-hint בלבד!\n\n"
        )

    # ─── Build email hint text ───
    email_hint_text = ''
    if email_parts_hint:
        email_hint_text = (
            f"📧 **מספרי פריט מהמייל** (hint — אחד מהם כנראה ב-Title Block):\n"
            f"   {email_parts_hint}\n"
            f"   ⚠️ קרא מהתמונה! זה רק hint!\n\n"
        )
    
    messages = [
        {"role": "system", "content": prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": (
                    f" **קרא רק מהתמונה!** \n\n"
                    f" **משימה**: חלץ מידע מ-Title Block של RAFAEL\n\n"
                    f" **התמונה הזו היא Title Block מוגדל x3!**\n"
                    f"   הטקסט צריך להיות ברור וקריא\n\n"
                    f" **הוראות קריאה - תו אחר תו**:\n"
                    f"1. מצא את השדה \"P.N.\" (Part Number)\n"
                    f"   🔴 **CAT NO. ≠ P.N.!** CAT NO. נמצא בשורה העליונה ליד DIM/MATERIAL.\n"
                    f"   P.N. נמצא בשורה נפרדת מתחת ל-TITLE, ליד SHT/OF.\n"
                    f"   אם שניהם קיימים — קרא רק מ-P.N., התעלם מ-CAT NO.!\n"
                    f"2. **קרא כל תו בנפרד לאט**:\n"
                    f"    תו 1: ___\n"
                    f"    תו 2: ___\n"
                    f"    תו 3: ___\n"
                    f"    תו 4: ___\n"
                    f"    וכן הלאה...\n\n"
                    f"3. **שים לב להבדלים קריטיים**:\n"
                    f"    B (אות)  8 (ספרה) - B יש קווים ישרים, 8 עגול\n"
                    f"    E (אות)  3 (ספרה) - E פתוח לצד, 3 פתוח למעלה ולמטה\n"
                    f"    0 (אפס)  O (אות) - 0 יכול להיות עם קו באמצע\n"
                    f"    5 (ספרה)  S (אות) - 5 יש קו ישר למעלה\n\n"
                    f"4. מצא \"DRAWING NO.\" - **קרא גם תו-אחר-תו**\n"
                    f"5. מצא \"REV\" - האות/מספר ליד\n\n"
                    f" 🔴🔴🔴 מקרה קריטי — P.N. ו-DRAWING NO. הם ערכים שונים לחלוטין!\n"
                    f"   דוגמה:\n"
                    f"     P.N.:        MMA76J800203B    ← זה מספר הפריט!\n"
                    f"     DRAWING NO.: 22-A19750        ← זה מספר השרטוט!\n"
                    f"   אם תחזיר part_number='22-A19750' — זו טעות חמורה!\n"
                    f"   קרא כל שדה בנפרד מהתא שלו!\n\n"
                    f" **דוגמאות מספרים אמיתיים לדקדוק**:\n"
                    f" BBLE4352A: B(אות) B(אות) L(אל) E(אות) 4 3 5 2 A\n"
                    f" HLTA10872A: H L T A 1 0 8 7 2 A (מ-P.N., לא 510030054 שזה CAT.NO!)\n"
                    f" 8H-A49880: 8(ספרה) H(אות) - A 4 9 8 8 0\n"
                    f" RF21MP614: R F 2 1 M P 6 1 4\n"
                    f" BG58498A: B G 5 8 4 9 8 A\n\n"
                    f" **OCR (לא אמין!)**:\n"
                    f"{'='*60}\n"
                    f"{text_snippet[:400] if text_snippet and text_snippet.strip() else '[התעלם - קרא מהתמונה!]'}\n"
                    f"{'='*60}\n\n"
                    f"📋 **pdfplumber text (מטקסט הPDF — אמין למספרים)**:\n"
                    f"{'='*60}\n"
                    f"{pdfplumber_hint if pdfplumber_hint else '[לא זמין]'}\n"
                    f"{'='*60}\n"
                    f"⚠️ אם pdfplumber מראה P.N. שונה מ-DRAWING NO. — שים לב! הם באמת שונים!\n"
                    f"   אמת תמיד מול התמונה, אבל pdfplumber אמין למספרים.\n\n"
                    f"{email_hint_text}"
                    f"{filename_hint_text}"
                    f" **עכשיו הסתכל בתמונה המוגדלת וקרא תו-אחר-תו!**\n"
                    f" Title Block location: {tb_location}"
                )},
                {"type": "image_url", "image_url": {
                    "url": f"data:image/png;base64,{image_to_use}",
                    "detail": "high"
                }}
            ]
        }
    ]
    
    # Call Vision API with automatic content filter retry
    response = _call_vision_api_with_retry(client, messages, max_tokens=350, temperature=0, stage_num=STAGE_BASIC_INFO)
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
            return None, 0, 0
    
    # Force customer_name to RAFAEL
    data['customer_name'] = 'RAFAEL'
    
    # Mark that we explicitly searched for P.N. field (so don't use drawing_number as fallback later)
    data['_searched_for_pn_field'] = True
    
    # ✅ VALIDATION: Ensure part_number is different from drawing_number
    # If they're the same AND extracted from explicit P.N. field - both are valid
    # But if part_number is empty or None, clear it (don't use drawing_number as fallback here)
    part_number = data.get('part_number', '').strip() if data.get('part_number') else ''
    drawing_number = data.get('drawing_number', '').strip() if data.get('drawing_number') else ''
    
    # If part_number is empty, leave it empty (will be handled later in post-processing)
    if not part_number and drawing_number:
        logger.warning(f"⚠️ WARNING Stage 1: part_number is empty, leaving blank (will not use drawing_number as fallback)")
        data['part_number'] = ''
    
    # If both exist and are identical, that's OK (P.N. and DRAWING NO. are sometimes the same)
    if part_number and drawing_number and part_number == drawing_number:
        logger.info(f"✓ Stage 1: P.N. and DRAWING NO. are identical: {part_number}")
    
    # If different, verify they're truly different fields
    if part_number and drawing_number and part_number != drawing_number:
        logger.info(f"✓ Stage 1: P.N.={part_number}, DRAWING NO.={drawing_number} (correctly different)")
    
    return data, usage.prompt_tokens, usage.completion_tokens


def extract_processes_info_rafael(text_snippet: str, image_b64: str, client: AzureOpenAI, **kwargs) -> Tuple[Optional[Dict], int, int]:
    """
    שלב 2 לרפאל: תהליכים ומפרטים - גרסת רפאל
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי לשנות את הפרומפט כאן
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return extract_processes_info(text_snippet, image_b64, client, **kwargs)


def extract_processes_from_notes(notes_text: str, client: AzureOpenAI) -> Tuple[Optional[Dict], int, int]:
    """
    שלב 5: Fallback - חילוץ תהליכים מטקסט NOTES בלבד (ללא תמונה)
    
    נעשה שימוש כשstage 2 (image-based) נכשל.
    שלב 3 כבר חילץ את כל ההנחיות - פה אנחנו מנתחים אותם לתהליכים.
    משתמשים בـ**אותו prompt** כמו Stage 2, רק text-only.
    """
    if not notes_text or not notes_text.strip():
        return None, 0, 0
    
    prompt = load_prompt("06b_extract_processes_from_notes")
    
    messages = [
        {"role": "system", "content": prompt},
        {
            "role": "user",
            "content": f"NOTES TEXT (from Stage 3 extraction):\n{notes_text}\n\nחלץ תהליכים ומפרטים מהטקסט - אותן שאלות כמו Stage 2."
        }
    ]
    
    # Call API without image (text only)
    try:
        stage5_model, stage5_max_tokens, stage5_temperature = _resolve_stage_call_config(
            STAGE_VALIDATION,
            400,
            0,
        )
        response = _chat_create_with_token_compat(
            client,
            model=stage5_model,
            messages=messages,
            max_tokens=stage5_max_tokens,
            temperature=stage5_temperature,
            response_format=RESPONSE_FORMAT
        )
    except Exception as e:
        logger.error(f"Stage 5: API error: {str(e)[:100]}")
        return None, 0, 0
    
    if response is None:
        return None, 0, 0
    
    content = response.choices[0].message.content
    usage = response.usage
    
    # Parse JSON
    try:
        data = json.loads(content)
    except json.JSONDecodeError as e:
        
        # Try to extract fields manually from the broken JSON
        data = {}
        
        # Extract fields using regex patterns
        fields = ['material', 'coating_processes', 'painting_processes', 'colors', 'marking_process', 
                  'part_dimensions', 'specifications', 'parts_list_page', 'process_summary_hebrew', 'process_summary_hebrew_short']
        
        for field in fields:
            # Pattern: "field_name": "value" - handle multiline and special chars
            pattern = rf'"{field}"\s*:\s*"([^"]*(?:\\.[^"]*)*)"'
            match = re.search(pattern, content, re.DOTALL)
            if match:
                try:
                    # Unescape the captured value
                    value = match.group(1)
                    # Handle escaped quotes and newlines
                    value = value.replace('\\"', '"').replace('\\n', ' ')
                    data[field] = value
                except Exception as e:
                    data[field] = None
            else:
                # Try alternative pattern for null values
                pattern_null = rf'"{field}"\s*:\s*null'
                if re.search(pattern_null, content):
                    data[field] = None
        
        if not data or all(v is None for v in data.values()):
            return None, 0, 0
    
    return data, usage.prompt_tokens, usage.completion_tokens


def extract_notes_text_rafael(text_snippet: str, image_b64: str, client: AzureOpenAI, **kwargs) -> Tuple[Optional[str], int, int]:
    """
    שלב 3 לרפאל: הנחיות מלאות - גרסת רפאל
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי לשנות את הפרומפט כאן
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return extract_notes_text(text_snippet, image_b64, client, **kwargs)


def extract_area_info_rafael(text_snippet: str, image_b64: str, client: AzureOpenAI) -> Tuple[Optional[str], int, int]:
    """
    שלב 4 לרפאל: חישוב שטח - גרסת רפאל
    כרגע זהה לגרסה הרגילה - בעתיד תוכלי לשנות את הפרומפט כאן
    """
    #  כרגע קורא לפונקציה הרגילה - בעתיד תשני את הפרומפט כאן
    return calculate_geometric_area(text_snippet, image_b64, client)
