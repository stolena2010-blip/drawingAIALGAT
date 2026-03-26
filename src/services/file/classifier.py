"""
File Type Classifier – Vision-API-based document classification
===============================================================
Extracted from customer_extractor_v3_dual.py  (Phase 2.6)
"""

import base64
import io
import json
import logging
import re
from pathlib import Path
from typing import Tuple

import pdfplumber
from openai import AzureOpenAI

from src.core.constants import (
    DRAWING_EXTS,
    STAGE_CLASSIFICATION,
    MAX_IMAGE_DIMENSION,
    WARN_IMAGE_DIMENSION,
)
from src.services.image.processing import _downsample_high_res_image
from src.services.ai.vision_api import _call_vision_api_with_retry
from src.services.file.file_utils import _detect_text_heavy_pdf
from src.utils.prompt_loader import load_prompt


# ── public API ──────────────────────────────────────────────────────────
__all__ = [
    "classify_file_type",
]


# ────────────────────────────────────────────────────────────────────────
def classify_file_type(file_path: str, client: AzureOpenAI) -> Tuple[str, str, str, str, int, int]:
    """
    זיהוי סוג קובץ - שרטוט או סוג אחר
    
    Returns:
        (file_type, description, quote_number, order_number, input_tokens, output_tokens)
        
    file_type options:
        - "DRAWING" - שרטוט הנדסי
        - "PURCHASE_ORDER" - הזמנת רכש
        - "QUOTE" - הצעת מחיר
        - "INVOICE" - חשבונית/הזמנה
        - "PARTS_LIST" - רשימת חלקים (PL)
        - "3D_MODEL" - קובץ תלת מימד (STEP, IGES, etc.)
        - "3D_IMAGE" - תמונת תלת מימד/ויזואליזציה
        - "ARCHIVE" - קובץ ארכיון (ZIP, RAR) - יחולץ אוטומטית
        - "OTHER" - אחר
    """
    file_path = Path(file_path)
    _is_likely_drawing = False
    logger = logging.getLogger(__name__)
    
    # Check file extension first
    ext = file_path.suffix.lower()
    filename_lower = file_path.stem.lower()
    
    # Archive files - will be extracted in Phase 0.5
    if ext in {'.zip', '.rar'}:
        return "ARCHIVE", f"Archive file ({ext}) - will be extracted automatically", "", "", 0, 0
    
    # 3D model files - no need for Vision API
    # Solidworks, Inventor, AutoCAD, CATIA, Fusion 360, Revit, 3D printing, and other CAD formats
    if ext in {
        # Solidworks
        '.sldprt', '.sldasm', '.slddrw',
        # Inventor
        '.ipt', '.iam', '.idw',
        # Fusion 360
        '.f3d', '.f3z',
        # Standard formats
        '.step', '.stp', '.iges', '.igs',
        # Autodesk formats
        '.x_t', '.x_b', '.sat', '.jt',
        # 3D printing formats
        '.3mf', '.stl', '.amf',
        # Generic 3D formats
        '.obj', '.ply', '.fbx', '.vrml', '.wrl', '.u3d', '.3ds',
        # CAD formats
        '.dwg', '.dxf', '.cgr',
        # Other CAD systems
        '.eprt', '.easm',
        '.catpart', '.catproduct', '.catdrawing',
        '.prt', '.asm', '.drw',
        # Revit/BIM
        '.rvt', '.nwd', '.nwc'
    }:
        return "3D_MODEL", f"3D model file ({ext})", "", "", 0, 0
    
    # Check filename for specific type indicators (before Vision API)
    # Check for PL (Parts List) - can be standalone, prefix like "PL1093Y815", 
    # suffix like "778PL", or compound like "PARTLIST"
    # Replace _ with - for PL detection (since _ is \w in regex, \b won't match)
    fn_for_pl = filename_lower.replace('_', '-')
    if re.search(r'(?<![a-z])pl(?![a-z])', fn_for_pl) or \
       re.search(r'(?<![a-z])pl\d', fn_for_pl) or \
       re.search(r'\d+pl(?![a-z])', fn_for_pl) or \
       re.search(r'partlist', filename_lower) or re.search(r'parts[\s_-]*list', filename_lower) or \
       "bom" in filename_lower:
        return "PARTS_LIST", "Parts list (identified by filename)", "", "", 0, 0
    
    if "model" in filename_lower or "asm" in filename_lower or "assembly" in filename_lower:
        return "3D_MODEL", "3D model (identified by filename)", "", "", 0, 0
    
    # 3D image files - תמונות מוצרים וויזואליזציות (stage ראשוני - לא ממתינים ל-AI)
    # כל תמונה בסיומות תמונה שלא היא מודל או PL - זוהה כ-3D_IMAGE מידי
    if ext in {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif', '.webp'}:
        return "3D_IMAGE", f"Product image / 3D visualization ({ext})", "", "", 0, 0
    
    # Only process visual files with Vision API (drawings mainly)
    if ext not in DRAWING_EXTS:
        return "OTHER", f"Unsupported format ({ext})", "", "", 0, 0

    # Guardrails specific to PDFs
    if ext == ".pdf":
        # Limit: if PDF has more than 10 pages, treat as document
        try:
            with pdfplumber.open(file_path) as pdf:
                page_count = len(pdf.pages)
                if page_count > 10:
                    return "OTHER", f"PDF עם {page_count} עמודים - מסווג כמסמך (מגבלת 10 עמודים לזיהוי שרטוט)", "", "", 0, 0

                # Aggregate text load across all pages to detect heavy narrative docs
                text_words = 0
                text_lines = 0
                _all_text = ""
                for p in pdf.pages:
                    txt = p.extract_text() or ""
                    _all_text += txt + "\n"
                    if txt:
                        text_words += len(txt.split())
                        text_lines += txt.count("\n") + 1
                
                # Check for engineering drawing keywords — if found, skip text-heavy check
                _all_text_upper = _all_text.upper()
                
                _drawing_keywords = [
                    'DRAWING NO', 'DWG NO', 'P.N.', 'PART NO',
                    'SCALE', 'REVISION', 'TOLERAN', 'SURFACE TEXTURE',
                    'PRODUCTION ROUTING', 'ROUTING CHART',
                    'MATERIAL', 'MACHINING', 'SURFACE TREATMENT',
                    'INSPECTION', 'PACKING', 'CAGE CODE',
                    'SECTION A-A', 'DETAIL', 'GD&T',
                ]
                _company_keywords = [
                    'RAFAEL', 'ELBIT', 'IAI', 'ISRAEL AEROSPACE',
                    'IMI', 'TADA', 'BIRD AERO',
                ]
                
                _kw_hits = sum(1 for kw in _drawing_keywords if kw in _all_text_upper)
                _company_hits = sum(1 for kw in _company_keywords if kw in _all_text_upper)
                _is_likely_drawing = (_kw_hits >= 3) or (_kw_hits >= 2 and _company_hits >= 1)
                
                if _is_likely_drawing:
                    logger.debug(f"Drawing keywords found ({_kw_hits} keywords, {_company_hits} company) — skipping text-heavy check")
                
                # If too much text overall — but scale threshold by page count
                # Multi-page drawings (A0 with 5-6 sheets) naturally have lots of text
                words_per_page = text_words / max(page_count, 1)
                lines_per_page = text_lines / max(page_count, 1)
                
                # Only flag as text-heavy if AVERAGE per page is high
                # Engineering drawings with routing charts/BOM can reach 500-600 words/page
                # Real text documents (contracts, specs) have 700+ words/page
                if words_per_page >= 700 and lines_per_page >= 100 and not _is_likely_drawing:
                    return "OTHER", (
                        f"מסמך מרובה מלל ({page_count} עמודים, ~{text_words} מילים, "
                        f"ממוצע {words_per_page:.0f} מילים/עמוד)"
                    ), "", "", 0, 0
        except Exception as e:
            print(f"      PDF page/text guard skipped: {e}")

        # Existing text-heavy first-page guard
        is_text_heavy, reason = _detect_text_heavy_pdf(file_path)
        if is_text_heavy and not _is_likely_drawing:
            return "OTHER", f"{reason} - סיווג כ-DOCUMENT/OTHER כדי למנוע False Positive כשרטוט", "", "", 0, 0
    
    try:
        # Get first page image
        img_bytes = None
        
        if ext == '.pdf':
            try:
                with pdfplumber.open(file_path) as pdf:
                    if not pdf.pages:
                        # PDF has no pages readable by pdfplumber - might be scanned/image-only
                        # Try to read it as image using different method
                        print(f"      PDF has no readable pages, trying alternative method...")
                        try:
                            # Try using pdf2image or PyMuPDF
                            import fitz  # PyMuPDF
                            doc = fitz.open(file_path)
                            if len(doc) > 0:
                                page = doc[0]
                                pix = page.get_pixmap(dpi=150)
                                img_bytes = pix.tobytes("png")
                                doc.close()
                                print(f"     Extracted image using PyMuPDF")
                            else:
                                return "OTHER", "PDF with no pages", "", "", 0, 0
                        except ImportError:
                            # PyMuPDF not available, try pdf2image
                            try:
                                from pdf2image import convert_from_path
                                images = convert_from_path(str(file_path), first_page=1, last_page=1, dpi=150)
                                if images:
                                    img_buffer = io.BytesIO()
                                    images[0].save(img_buffer, format='PNG')
                                    img_bytes = img_buffer.getvalue()
                                    print(f"     Extracted image using pdf2image")
                                else:
                                    return "OTHER", "Failed to extract image from PDF", "", "", 0, 0
                            except Exception as e:
                                return "OTHER", f"PDF read error: {str(e)}", "", "", 0, 0
                    else:
                        # Normal PDF with readable pages
                        page = pdf.pages[0]
                        img = page.to_image(resolution=150)
                        
                        img_buffer = io.BytesIO()
                        img.original.save(img_buffer, format='PNG')
                        img_bytes = img_buffer.getvalue()
            except Exception as e:
                return "OTHER", f"PDF processing error: {str(e)}", "", "", 0, 0
        else:
            # Image file
            with open(file_path, 'rb') as f:
                img_bytes = f.read()
        
        if not img_bytes:
            return "OTHER", "Failed to read file", "", "", 0, 0
        
        # Check and downsample high resolution images
        img_bytes, was_downsampled, orig_size, new_size = _downsample_high_res_image(img_bytes, MAX_IMAGE_DIMENSION)
        
        if was_downsampled:
            print(f"     Downsampled: {orig_size[0]}x{orig_size[1]}  {new_size[0]}x{new_size[1]}")
        elif max(orig_size) > WARN_IMAGE_DIMENSION:
            print(f"      High resolution: {orig_size[0]}x{orig_size[1]}")
        
        # Encode to base64
        image_b64 = base64.b64encode(img_bytes).decode('utf-8')
        
        # Vision API prompt - MAXIMUM ACCURACY
        prompt = load_prompt("09_classify_document_type")
        
        messages = [
            {"role": "system", "content": prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": " סרוק את המסמך בקפידה וזהה את סוגו בדיוק מקסימלי. חשוב במיוחד לזהות שרטוטים הנדסיים נכון!"},
                    {"type": "image_url", "image_url": {
                        "url": f"data:image/png;base64,{image_b64}",
                        "detail": "low"  # LOW detail sufficient for file type classification
                    }}
                ]
            }
        ]
        
        # Call Vision API with automatic content filter retry
        response = _call_vision_api_with_retry(client, messages, max_tokens=200, temperature=0, stage_num=STAGE_CLASSIFICATION)
        if response is None:
            return "OTHER", "Vision API failed after retry", "", "", 0, 0
        
        content_raw = response.choices[0].message.content
        if isinstance(content_raw, list):
            text_parts = []
            for part in content_raw:
                if isinstance(part, dict):
                    text_value = part.get("text")
                    if text_value:
                        text_parts.append(str(text_value))
                elif part:
                    text_parts.append(str(part))
            content = "\n".join(text_parts)
        else:
            content = str(content_raw or "")
        usage = response.usage
        
        # Parse JSON
        try:
            data = json.loads(content)
        except json.JSONDecodeError:
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group())
            else:
                content_upper = content.upper()
                detected_type = "OTHER"
                type_patterns = [
                    ("PURCHASE_ORDER", r"\bPURCHASE[ _-]?ORDER\b|\bPO\b"),
                    ("PARTS_LIST", r"\bPARTS[ _-]?LIST\b|\bBOM\b|\bBILL OF MATERIALS\b"),
                    ("INVOICE", r"\bINVOICE\b|\bTAX INVOICE\b"),
                    ("QUOTE", r"\bQUOTE\b|\bQUOTATION\b|\bRFQ\b|\bPROPOSAL\b"),
                    ("DRAWING", r"\bDRAWING\b|\bDWG\b|\bTITLE BLOCK\b"),
                    ("OTHER", r"\bOTHER\b|\bDOCUMENT\b"),
                ]
                for candidate, pattern in type_patterns:
                    if re.search(pattern, content_upper, re.IGNORECASE):
                        detected_type = candidate
                        break

                confidence = "low"
                if re.search(r"\bHIGH\b", content_upper):
                    confidence = "high"
                elif re.search(r"\bMEDIUM\b", content_upper):
                    confidence = "medium"

                desc = (content or "").strip()
                if len(desc) > 250:
                    desc = desc[:250] + "..."

                data = {
                    "file_type": detected_type,
                    "description": desc or "Parsed from free-text model response",
                    "confidence": confidence,
                    "quote_number": "",
                    "order_number": "",
                }
        
        file_type = data.get('file_type', 'OTHER')
        description = data.get('description', '')
        confidence = data.get('confidence', 'low')
        quote_number = data.get('quote_number', '')
        order_number = data.get('order_number', '')
        
        full_desc = f"{description} (confidence: {confidence})"
        
        return file_type, full_desc, quote_number, order_number, usage.prompt_tokens, usage.completion_tokens
        
    except Exception as e:
        print(f"    ** Classification failed: {e}")
        return "OTHER", f"Classification error: {str(e)}", "", "", 0, 0
