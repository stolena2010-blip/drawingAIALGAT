"""
B2B Text Export – save TAB-delimited summary files
===================================================
Extracted from customer_extractor_v3_dual.py  (Phase 2.5)
"""

from pathlib import Path
from typing import List, Dict
import re

from src.utils.logger import get_logger
logger = get_logger(__name__)


def _is_single_numeric_quantity(value: str) -> bool:
    """
    Check if a quantity string is a single clean number suitable for B2B field 4.
    Returns True for: '50', '153', '1.0000', '4500'
    Returns False for: '(80, 100)', '10-20', '3, 5, 10', '10-20 יחידות',
                        '0', text with numbers, ranges, multiple quantities
    """
    if not value or value == '0':
        return False
    # Strip parentheses
    stripped = value.strip().strip('()')
    # Must be a single number (possibly with decimal point)
    return bool(re.match(r'^\d+(\.\d+)?$', stripped))


# ── public API ──────────────────────────────────────────────────────────
__all__ = [
    "_save_text_summary",
    "_save_text_summary_with_variants",
]


# ────────────────────────────────────────────────────────────────────────
def _save_text_summary(results: List[Dict], output_path: Path, customer_email: str = "", b2b_number: str = "", timestamp: str = "") -> None:
    """
    שמירת קובץ טקסט מסכם עם TAB delimiters
    
    Args:
        results: רשימת תוצאות (פריטים)
        output_path: נתיב לקובץ טקסט
        customer_email: כתובת מייל של הלקוח
        b2b_number: מספר B2B
        timestamp: חותמת זמן
    """
    try:
        rows = []
        row_number = 0
        
        
        for item in results:
            # קבל מספר פריט
            part_num = str(item.get('part_number', '')).strip()
            
            
            # דלג רק על פריטים ממש ריקים
            if not part_num:
                logger.debug(f"Skipping (no part_number)")
                continue
            
            row_number += 1
            
            # פונקציה עזר להמיר ערכים ריקים ל-string ריק
            def safe_str(value, default=''):
                """המר ערך ל-string, אם None או ריק החזר default"""
                if value is None or value == '':
                    return default
                return str(value).strip()
            
            # בנה שדה 11 (תיאור מורחב) - merged_description, fallback לסיכום בעברית
            hebrew_desc = safe_str(item.get('merged_description', ''))
            if not hebrew_desc:
                hebrew_desc = safe_str(item.get('process_summary_hebrew', ''))
            if not hebrew_desc:
                # Fallback: merged_specs or merged_notes or item_name
                hebrew_desc = safe_str(item.get('merged_specs', ''))
            if not hebrew_desc:
                hebrew_desc = safe_str(item.get('merged_notes', ''))
            if not hebrew_desc:
                hebrew_desc = safe_str(item.get('item_name', ''))
            
            # בנה שדה 4 (כמות) ושדה 9 (הערות)
            quantity_value = safe_str(item.get('quantity'), '1.0000')
            notes_field = ''  # שדה 9 - הערות
            
            # כמות מספרית יחידה (לדוגמה "153") → שדה 4
            # כל השאר (טווח, טקסט, כמויות מרובות, סוגריים) → שדה 9
            if _is_single_numeric_quantity(quantity_value):
                quantity_field = quantity_value.strip().strip('()')
                notes_field = ''
            else:
                quantity_field = '0'
                notes_field = quantity_value if quantity_value != '1.0000' else ''
            
            # בנה שדה 10 (רמת ביטחון) - confidence level
            confidence = safe_str(item.get('confidence_level', ''))
            
            # בנה שדה 17 - גרסת פריט בלבד
            revision = safe_str(item.get('revision'), '')
            field_17 = revision
            
            # בנה שורה עם 17 שדות מופרדים ב-TAB (כל שדה תמיד קיים, גם אם ריק)
            fields = [
                str(row_number),  # 1: מספר שדה
                part_num,  # 2: מספר שורת חומרה
                revision,  # 3: גרסת פריט
                quantity_field,  # 4: כמות (0 אם יש כמה כמויות)
                '1',  # 5: יחידה מידה
                '0.0000',  # 6: מחיר יחידה
                '0',  # 7: מטבע הזמנה
                safe_str(item.get('delivery_date')),  # 8: תאריך אספקה מבוקש
                notes_field,  # 9: הערות (כמויות מרובות אם קיימות)
                confidence,  # 10: רמת ביטחון
                hebrew_desc,  # 11: תיאור מורחב (merged_description)
                safe_str(item.get('item_name')),  # 12: תיאור פריט/עבודה באנגלית
                '0',  # 13: מספר B2B (תמיד 0)
                timestamp or '',  # 14: מספר בקשה להצעת מחיר / TIMESTAMP
                customer_email or '',  # 15: מייל לקוח
                safe_str(item.get('drawing_number')),  # 16: מספר שרטוט
                field_17,  # 17: גרסת פריט בלבד
                safe_str(item.get('part_number_ocr_original'))  # 18: מספר פריט OCR מקורי (אם הוחלף ע"י PL)
            ]
            
            # הצטרף עם TAB
            row = '\t'.join(fields)
            logger.debug(f"Row {row_number}: {fields[0]}\t{fields[1]}\t{fields[2]}\t{fields[3]}")
            rows.append(row)
        
        # שמור לקובץ עם separator מיוחד בין שורות
        # כל שורה מסתיימת ב-{~#~} ולאחריה ירידת שורה
        content = '{~#~}\n'.join(rows)
        # הוסף {~#~} גם בסוף הקובץ (כמו לאחרי כל שורה)
        if content:
            content += '{~#~}\n'
        
        # שמור עם CP1255 (Windows-1255) encoding - encoding ישראלי סטנדרטי
        # CP1255 תומך בעברית ומוכר למערכות וינדוס ומסדי נתונים בישראל
        with open(output_path, 'w', encoding='cp1255', errors='replace') as f:
            f.write(content)
        
        # בדוק שהקובץ נכתב
        file_size = output_path.stat().st_size
        logger.info(f"✓ Summary text file saved: {output_path.name} ({len(rows)} rows, {file_size} bytes)")
        
    except Exception as e:
        logger.error(f"Failed to save text file: {e}")
        import traceback
        traceback.print_exc()
        logger.debug(f"{len(rows)} items")
        
    except Exception as e:
        logger.error(f"Failed to save text file: {e}")


# ────────────────────────────────────────────────────────────────────────
def _save_text_summary_with_variants(results: List[Dict], output_path: Path, customer_email: str = "", b2b_number: str = "", timestamp: str = "") -> None:
    """
    שמירת שלושה קבצי טקסט מסכם עם סינון רמת ביטחון:
    - B2B-... כל השורות
    - B2BH-... HIGH + FULL בלבד
    - B2BM-... MEDIUM + HIGH + FULL
    """
    try:
        def safe_str(value, default=''):
            if value is None or value == '':
                return default
            return str(value).strip()
        
        def build_rows(results, confidence_filter=None):
            rows = []
            row_number = 0
            
            for item in results:
                part_num = str(item.get('part_number', '')).strip()
                if not part_num:
                    continue
                
                if confidence_filter:
                    conf_level = safe_str(item.get('confidence_level', '')).upper()
                    if conf_level == 'FULL':
                        conf_level = 'HIGH'
                    if conf_level not in confidence_filter:
                        continue
                
                row_number += 1
                hebrew_desc = safe_str(item.get('merged_description', ''))
                if not hebrew_desc:
                    hebrew_desc = safe_str(item.get('process_summary_hebrew', ''))
                if not hebrew_desc:
                    hebrew_desc = safe_str(item.get('merged_specs', ''))
                if not hebrew_desc:
                    hebrew_desc = safe_str(item.get('merged_notes', ''))
                if not hebrew_desc:
                    hebrew_desc = safe_str(item.get('item_name', ''))
                quantity_value = safe_str(item.get('quantity'), '1.0000')
                notes_field = ''
                
                # כמות מספרית יחידה (לדוגמה "153") → שדה 4
                # כל השאר (טווח, טקסט, כמויות מרובות, סוגריים) → שדה 9
                if _is_single_numeric_quantity(quantity_value):
                    quantity_field = quantity_value.strip().strip('()')
                    notes_field = ''
                else:
                    quantity_field = '0'
                    notes_field = quantity_value if quantity_value != '1.0000' else ''
                
                confidence = safe_str(item.get('confidence_level', ''))
                revision = safe_str(item.get('revision'), '')
                
                fields = [
                    str(row_number),
                    part_num,
                    revision,
                    quantity_field,
                    '1',
                    '0.0000',
                    '0',
                    safe_str(item.get('delivery_date')),
                    notes_field,
                    confidence,
                    hebrew_desc,
                    safe_str(item.get('item_name')),
                    '0',
                    timestamp or '',
                    customer_email or '',
                    safe_str(item.get('drawing_number')),
                    revision
                ]
                
                row = '\t'.join(fields)
                rows.append(row)
            
            return rows
        
        def save_variant(variant_letter, rows):
            if not rows:
                return None
            
            base_name = output_path.stem
            if variant_letter:
                base_parts = base_name.split('-')
                variant_filename = f"{base_parts[0]}{variant_letter}-{'-'.join(base_parts[1:])}.txt"
            else:
                variant_filename = f"{base_name}.txt"
            
            variant_path = output_path.parent / variant_filename
            
            content = '{~#~}\n'.join(rows)
            if content:
                content += '{~#~}\n'
            
            with open(variant_path, 'w', encoding='cp1255', errors='replace') as f:
                f.write(content)
            
            file_size = variant_path.stat().st_size
            variant_label = 'B2BH' if variant_letter == 'H' else 'B2BM' if variant_letter == 'M' else 'B2B'
            logger.info(f"✓ {variant_label} file saved: {variant_path.name} ({len(rows)} rows, {file_size} bytes)")
            return variant_path
        
        all_rows = build_rows(results)
        high_rows = build_rows(results, confidence_filter={'HIGH'})
        medium_high_rows = build_rows(results, confidence_filter={'MEDIUM', 'HIGH'})
        
        logger.info(f"Creating 3 B2B text file variants...")
        save_variant('', all_rows)
        save_variant('H', high_rows)
        save_variant('M', medium_high_rows)
        
    except Exception as e:
        logger.error(f"Failed to save text file variants: {e}")
        import traceback
        traceback.print_exc()
