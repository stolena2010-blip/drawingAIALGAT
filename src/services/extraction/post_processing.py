"""
Post-processing for extraction results.
=========================================

Extracted from customer_extractor_v3_dual.py for modularity.

Functions:
- post_process_summary_from_notes(): Ensure MASK/INSERTS from NOTES appear in summary
"""

import re

from src.utils.logger import get_logger

logger = get_logger(__name__)


def post_process_summary_from_notes(result_data: dict, pdfplumber_text: str = '', ocr_text: str = '') -> dict:
    """
    Post-processing דטרמיניסטי: מוודא שמיסוך וקשיחים מופיעים בסיכום
    כשהם קיימים ב-notes_full_text.

    Stage 2 (Vision API) לא דטרמיניסטי — לפעמים מפספס MASK/INSERTS
    למרות שהם מופיעים ב-NOTES. פונקציה זו מתקנת את הפער.

    נקרא אחרי merge של כל ה-stages, לפני P.N. extraction.
    """
    notes_text = str(result_data.get('notes_full_text') or '')
    notes_upper = notes_text.upper()

    if not notes_upper.strip():
        return result_data

    summary = result_data.get('process_summary_hebrew', '') or ''
    short = result_data.get('process_summary_hebrew_short', '') or ''
    coating = result_data.get('coating_processes', '') or ''

    changed = False

    # ─── 1. MASKING: Ensure מיסוך appears if MASK is in NOTES ───
    mask_match = re.search(
        r'MASK\s+(.*?)(?:BEFORE|PRIOR|DURING|$)',
        notes_upper,
        re.MULTILINE
    )
    if mask_match and 'מיסוך' not in summary:
        # Build Hebrew masking description
        raw = mask_match.group(1).strip().rstrip(',. ')
        translations = [
            ('HOLES', 'חורים'), ('THREADS', 'הברגות'),
            ('MARKED SURFACES', 'משטחים מסומנים'),
            ('SURFACES', 'משטחים'), ('AND', 'ו'),
        ]
        mask_heb = raw
        for eng, heb in translations:
            mask_heb = mask_heb.replace(eng, heb)
        mask_heb = re.sub(r'\s+', ' ', f"מיסוך {mask_heb}").strip()

        # Insert in correct position: after ציפוי (②), before סימון/חריטה (④)
        parts = [p.strip() for p in summary.split('|')]
        insert_idx = len(parts)
        for idx, part in enumerate(parts):
            if any(kw in part for kw in ['סימון', 'חריטה', 'מילוי', 'שימון', 'קשיחים']):
                insert_idx = idx
                break
        parts.insert(insert_idx, mask_heb)
        summary = ' | '.join(parts)

        # Also fix short summary
        if 'מיסוך' not in short:
            short_parts = [p.strip() for p in short.split('|')]
            s_idx = len(short_parts)
            for idx, part in enumerate(short_parts):
                if any(kw in part for kw in ['סימון', 'חריטה', 'מילוי', 'שימון', 'קשיחים']):
                    s_idx = idx
                    break
            short_parts.insert(s_idx, 'מיסוך')
            short = ' | '.join(short_parts)

        # Also fix coating_processes if MASK missing there
        if 'mask' not in coating.lower():
            mask_eng = mask_match.group(0).strip()
            coating = f"{coating}, {mask_eng}" if coating else mask_eng

        changed = True
        logger.info(f"🔧 POST-PROCESS: Added masking → {mask_heb}")

    # ─── 2. INSERTS: Ensure קשיחים appears if INSERTS in NOTES ───
    has_inserts_in_notes = bool(re.search(
        r'INSERT(?:S)?\s+INSTALLATION|INSTALL\s+INSERTS',
        notes_upper
    ))
    if has_inserts_in_notes and 'קשיחים' not in summary:
        # Try to find CAT NO and QTY from BOM table in pdfplumber/OCR text
        combined_raw = f"{pdfplumber_text}\n{ocr_text}"
        insert_info = "קשיחים"

        # Pattern: INSERT...description...CAT_NO(9+ digits)...QTY(1-4 digits)...EA
        bom_match = re.search(
            r'INSERT[^\n]*?(\d{9,})\s+(\d{1,4})\s+EA',
            combined_raw,
            re.IGNORECASE
        )
        if bom_match:
            cat_no = bom_match.group(1)
            qty = bom_match.group(2)
            insert_info = f"קשיחים: {cat_no}×{qty}"
        else:
            # Broader search: any 9+ digit number near INSERT
            bom_lines = [l for l in combined_raw.split('\n')
                         if 'INSERT' in l.upper()]
            for line in bom_lines:
                nums = re.findall(r'\b(\d{9,})\b', line)
                qtys = re.findall(r'\b(\d{1,3})\s*(?:EA|PCS|יח)', line, re.IGNORECASE)
                if nums:
                    cat_no = nums[0]
                    qty = qtys[0] if qtys else ''
                    insert_info = f"קשיחים: {cat_no}×{qty}" if qty else f"קשיחים: {cat_no}"
                    break

        # Append at end (position ⑥)
        summary = f"{summary} | {insert_info}" if summary else insert_info
        if 'קשיחים' not in short:
            short = f"{short} | קשיחים" if short else "קשיחים"

        changed = True
        logger.info(f"🔧 POST-PROCESS: Added inserts → {insert_info}")

    # ─── Apply changes ───
    if changed:
        result_data['process_summary_hebrew'] = summary
        result_data['process_summary_hebrew_short'] = short
        result_data['coating_processes'] = coating

    return result_data
