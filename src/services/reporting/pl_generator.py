"""
Parts-List (PL) Generator – Hebrew / English summaries + AI extraction
======================================================================
Extracted from customer_extractor_v3_dual.py  (Phase 2.5)
"""

import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional

from openai import AzureOpenAI

from src.core.constants import STAGE_PL
from src.services.ai.vision_api import (
    _resolve_stage_call_config,
    _chat_create_with_token_compat,
)
from src.services.extraction.filename_utils import _normalize_item_number
from src.services.extraction.insert_validator import validate_inserts_hardware
from src.services.extraction.insert_price_lookup import lookup_insert_price
from src.utils.logger import get_logger
from src.utils.prompt_loader import load_prompt

logger = get_logger(__name__)

# ── public API ──────────────────────────────────────────────────────────
__all__ = [
    "_generate_pl_summary_hebrew",
    "_generate_pl_summary_english",
    "extract_pl_data",
]


# ────────────────────────────────────────────────────────────────────────
def _generate_pl_summary_hebrew(part_number: str, pl_items: List[Dict], ocr_original: str = '') -> str:
    """
    Generate Hebrew summary of PL items for a specific part number.
    FILTERED: Shows ONLY specific processes and materials
    
    Processes included: צפוי, צביעה, בדיקה, איריזה, מיסוך
    Materials included: קשיחים (inserts), צבעים (colors)
    
    Args:
        part_number: Part number to match (from drawing)
        pl_items: List of PL item dictionaries (with associated_item field from file_classifications)
        ocr_original: Original OCR part number before PL override (if applicable)
    
    Returns:
        Hebrew summary string format: "תהליך - מפרט | תהליך - מפרט | צבע | קשיח"
    """
    if not pl_items or not part_number:
        return ""
    
    # Whitelist of allowed processes (Hebrew and English variations)
    ALLOWED_PROCESSES = {
        'expected', 'צפוי',
        'painting', 'צביעה', 'paint',
        'testing', 'בדיקה', 'ביקורת', 'inspection', 'check',
        'packing', 'אריזה', 'pack',
        'passivation', 'פסיבציה', 'passivate',
        'anodizing', 'אנודייז', 'anodize',
        'masking', 'מיסוך', 'mask',
        'marking', 'סימון', 'mark',
        'laser marking', 'סימון בלייזר', 'laser',
        'coating', 'ציפוי',
        'machining', 'עיבוד', 'machine'
    }
    
    # Normalize the part number for comparison
    part_num_norm = _normalize_item_number(str(part_number))
    ocr_orig_norm = _normalize_item_number(str(ocr_original)) if ocr_original else ''
    
    matched_items = []
    for pl_item in pl_items:
        # Match by associated_item (the drawing this PL is linked to)
        associated = pl_item.get('associated_item', '')
        associated_norm = _normalize_item_number(str(associated))
        
        # Check if this PL belongs to the current drawing
        matched = (associated_norm == part_num_norm or
            associated_norm in part_num_norm or
            part_num_norm in associated_norm)
        # Also match against original OCR PN (before PL override)
        if not matched and ocr_orig_norm:
            matched = (associated_norm == ocr_orig_norm or
                associated_norm in ocr_orig_norm or
                ocr_orig_norm in associated_norm)
        if matched:
            matched_items.append(pl_item)
    
    if not matched_items:
        return ""
    
    # Collect FILTERED process-spec pairs and materials
    process_spec_pairs = []  # List of (process, spec) tuples
    materials = set()   # Set of unique materials (inserts, colors)
    
    for item in matched_items:
        # Get processes (English) for filtering
        processes_eng = item.get('processes', [])
        if isinstance(processes_eng, str):
            processes_eng = [processes_eng] if processes_eng else []
        
        # Get Hebrew processes for display
        processes_heb = item.get('processes_hebrew', [])
        if isinstance(processes_heb, str):
            processes_heb = [processes_heb] if processes_heb else []
        
        # Get specifications
        specifications = item.get('specifications', [])
        if isinstance(specifications, str):
            specifications = [specifications] if specifications else []
        
        description = (item.get('description') or '').strip().lower()
        item_type = (item.get('item_type') or '').strip().lower()
        product_tree = (item.get('product_tree') or '').strip()
        
        # Pair processes with specs, filtering by allowed processes using ENGLISH versions
        if processes_eng and specifications:
            for i, process_eng in enumerate(processes_eng):
                # Check if this English process is allowed
                if any(allowed in process_eng.lower() for allowed in ALLOWED_PROCESSES):
                    spec = specifications[i] if i < len(specifications) else (specifications[0] if specifications else "")
                    # Use Hebrew version for display (or English if translation failed)
                    process_heb = processes_heb[i] if i < len(processes_heb) else process_eng
                    if spec:
                        pair = (process_heb, spec.strip())
                        if pair not in process_spec_pairs:
                            process_spec_pairs.append(pair)
                    else:
                        pair = (process_heb, "")
                        if pair not in process_spec_pairs:
                            process_spec_pairs.append(pair)
        elif processes_eng:
            for i, process_eng in enumerate(processes_eng):
                # Check if this English process is allowed
                if any(allowed in process_eng.lower() for allowed in ALLOWED_PROCESSES):
                    # Use Hebrew version for display
                    process_heb = processes_heb[i] if i < len(processes_heb) else process_eng
                    pair = (process_heb, "")
                    if pair not in process_spec_pairs:
                        process_spec_pairs.append(pair)
        
        # MATERIALS: Only extract inserts and colors from product_tree or item_type
        if product_tree:
            tree_lower = product_tree.lower()
            if 'insert' in tree_lower or 'fastener' in tree_lower:
                # Extract insert/fastener info
                materials.add(description) if description else None
            if 'color' in tree_lower:
                materials.add(description) if description else None
        
        # Also check item_type for materials
        if item_type:
            if 'color' in item_type or 'coil_insert' in item_type or 'fastener' in item_type:
                materials.add(description) if description else None
    
    # Build summary string
    summary_parts = []
    
    # Add filtered process-spec pairs
    for process, spec in process_spec_pairs:
        if process and spec:
            summary_parts.append(f"{process} - {spec}")
        elif process:
            summary_parts.append(process)
        elif spec:
            summary_parts.append(spec)
    
    # Add materials (inserts and colors) at the end
    for material in sorted(materials):
        if material and material.strip():
            summary_parts.append(material)
    
    return " | ".join(summary_parts) if summary_parts else ""


# ────────────────────────────────────────────────────────────────────────
# PL Main Part Number Extraction (deterministic, no API cost)
# ────────────────────────────────────────────────────────────────────────

# TY codes that indicate MANUFACTURED parts (parts we make)
_MANTRA_MANUFACTURED_TY = {'MP', 'CM', 'B'}
# B = Basic manufactured part (seen in IAI/Golan MANTRA PLs)
# TY codes that indicate assemblies (ignore for part number)
_MANTRA_ASSEMBLY_TY = {'BM', 'SM', 'MA'}
# TY codes that indicate BUY items (ignore)
_MANTRA_BUY_TY = {'HM', 'H', 'HR', 'SP', '2B', 'F'}
# KBM manufactured types
_KBM_MANUFACTURED_TY = {'IN', 'IM', 'KT', 'AM', 'S1', 'S2', 'S3'}
# AM = Assembly Make, S1/S2/S3 = Sub-assembly levels (seen in KBM PLs)

# Words that should NEVER be treated as part numbers (column headers, labels, etc.)
_NOT_PART_NUMBERS = {
    'category', 'description', 'part', 'number', 'rev', 'revision',
    'catalog', 'seq', 'item', 'type', 'make', 'buy', 'qty', 'quantity',
    'status', 'release', 'name', 'unit', 'level', 'alt', 'notes',
    'find', 'reference', 'drawing', 'date', 'weight', 'material',
    'process', 'spec', 'supplier', 'price', 'cost', 'total', 'uom',
    'none', 'null', 'n/a', 'na', 'tbd', 'multiple',
}


def _detect_pl_format(pl_text: str) -> str:
    """Detect PL format from extracted text.
    
    Returns one of: 'BOEING', 'MANTRA_GOLAN', 'MANTRA_KBM', 'MANTRA_SYSTEMS', 'NEW_TABULAR', 'UNKNOWN'
    """
    # Boeing PLs contain "Type:PL ID:" which falsely matches MANTRA_GOLAN.
    # Detect Boeing FIRST by its unique signatures.
    if 'Document Number:' in pl_text and 'Type:PL' in pl_text:
        return 'BOEING'
    if 'BOEING' in pl_text.upper() and 'PARTS LIST' in pl_text:
        return 'BOEING'
    
    if 'PL ID:' in pl_text or 'PL ID :' in pl_text:
        return 'MANTRA_GOLAN'
    if 'Doc ID' in pl_text and 'UCP-' in pl_text:
        return 'MANTRA_KBM'
    if 'Doc ID' in pl_text:
        return 'MANTRA_SYSTEMS'
    if ('Part Number' in pl_text and 'Catalog No' in pl_text) or 'Release Status' in pl_text:
        return 'NEW_TABULAR'
    return 'UNKNOWN'


def _extract_header_part_number(pl_text: str, pl_format: str) -> str:
    """Extract the header/document part number from PL text."""
    if pl_format == 'BOEING':
        # Boeing format: "Document Number:173W7200-703 Type:PL ID: REV:A"
        # The real part number is after "Document Number:"
        match = re.search(r'Document\s*Number\s*:\s*([A-Z0-9][\w-]{4,})', pl_text)
        if match:
            return match.group(1).strip()
        # Fallback: look for CAGE NUMBER line, P.N. is nearby
        match = re.search(r'CONTROL\s+([A-Z0-9]{3,}\w*-\d{3})', pl_text[:1500])
        if match:
            return match.group(1).strip()
        return ''

    if pl_format == 'MANTRA_GOLAN':
        match = re.search(r'PL\s*ID\s*:\s*(\S+)', pl_text)
        if match:
            candidate = match.group(1).strip()
            # Reject revision identifiers — "REV:A", "REV:B", "REV:-" etc.
            if candidate.upper().startswith('REV'):
                logger.warning(f"PL header P.N. '{candidate}' looks like revision, not P.N. — skipping")
                return ''
            return candidate
        return ''
    
    if pl_format in ('MANTRA_SYSTEMS', 'MANTRA_KBM'):
        match = re.search(r'Doc\s*ID\s*[:\s]+\s*(\S+)', pl_text)
        return match.group(1).strip() if match else ''
    
    if pl_format == 'NEW_TABULAR':
        # Strategy 1: "Part Number  Rev" header row, then value on next line
        match = re.search(r'Part\s+Number\s+Rev\b', pl_text)
        if match:
            after = pl_text[match.end():match.end() + 300]
            pn_match = re.search(r'([A-Z0-9][A-Z0-9]{2,}-\d{3})', after)
            if pn_match:
                return pn_match.group(1)
        
        # Strategy 2: "Part Number" alone (columns may be on separate lines)
        # then look for -XXX pattern within nearby text
        match = re.search(r'Part\s+Number\b', pl_text)
        if match:
            after = pl_text[match.end():match.end() + 300]
            pn_match = re.search(r'([A-Z0-9][A-Z0-9]{2,}-\d{3})', after)
            if pn_match:
                return pn_match.group(1)
        
        # Strategy 3: ELTA-style part numbers near the top
        first_800 = pl_text[:800]
        pn_match = re.search(r'(\d{4}[A-Z]\d{3}-\d{3})', first_800)
        if pn_match:
            return pn_match.group(1)
        
        # Strategy 4: Any alphanumeric-XXX pattern (4+ chars before dash) in first 800
        pn_match = re.search(r'\b([A-Z0-9]{4,}-\d{3})\b', first_800)
        if pn_match:
            return pn_match.group(1)
    
    return ''


def _has_suffix(part_number: str) -> bool:
    """Check if part number already has a dash-number suffix like -001, -002."""
    return bool(re.search(r'-\d{3}$', part_number))


def _extract_manufactured_items_from_text(pl_text: str, pl_format: str) -> list:
    """Extract manufactured item part numbers from PL text using regex.
    
    Returns list of manufactured part number strings.
    """
    manufactured = []
    
    if pl_format in ('MANTRA_GOLAN', 'MANTRA_SYSTEMS'):
        # Pattern: ITM  PART_NUMBER  TY  PART_NAME
        # Lines with digit ITM, then part number, then 2-letter TY code
        for match in re.finditer(
            r'^\s*(\d{1,3})\s+(\S+)\s+(\w{1,3})\s+',
            pl_text, re.MULTILINE
        ):
            itm_num = int(match.group(1))
            part_num = match.group(2)
            ty_code = match.group(3).upper()
            
            # Skip raw materials (ITM 901+)
            if itm_num >= 901:
                continue
            # Skip assembly markers (items starting with *)
            if part_num.startswith('*'):
                continue
            # Skip ALT/GL_alt items
            if part_num.upper().startswith('ALT') or part_num.upper().startswith('GL_ALT'):
                continue
            
            if ty_code in _MANTRA_MANUFACTURED_TY:
                manufactured.append(part_num)
    
    elif pl_format == 'MANTRA_KBM':
        for match in re.finditer(
            r'^\s*(\d{1,3})\s+(\S+)\s+(\w{1,3})\s+(.*)',
            pl_text, re.MULTILINE
        ):
            itm_num = int(match.group(1))
            part_num = match.group(2)
            ty_code = match.group(3).upper()
            rest = match.group(4)
            
            if itm_num >= 901:
                continue
            if '*DELETED' in rest.upper() or '*DELETED' in part_num.upper():
                continue
            if part_num.upper().startswith('ALT') or part_num.upper().startswith('GL_ALT'):
                continue
            
            if ty_code in _KBM_MANUFACTURED_TY:
                manufactured.append(part_num)
    
    elif pl_format == 'NEW_TABULAR':
        # Look for "Make" items in the parts list
        # Pattern varies but typically: Seq No, Part Number, ..., Item Type
        for match in re.finditer(
            r'(\S+)\s+(\S+).*?\b(Make|Buy)\b',
            pl_text, re.IGNORECASE
        ):
            seq = match.group(1)
            part_num = match.group(2)
            item_type = match.group(3)
            
            # Skip if seq is 901+
            try:
                if int(seq) >= 901:
                    continue
            except ValueError:
                pass
            if seq.upper().startswith('ALT') or seq.upper().startswith('GL_ALT'):
                continue
            
            # Skip common column headers / non-PN words
            if part_num.lower() in _NOT_PART_NUMBERS:
                continue
            # Skip values that are too short or purely alphabetic without digits
            if len(part_num) < 3:
                continue
            
            if item_type.lower() == 'make':
                manufactured.append(part_num)
    
    return manufactured


def _determine_pl_main_part_number(pl_text: str) -> str:
    """Determine the main part number from a PL document.
    
    Logic:
    - Detect PL format
    - Extract header part number
    - Count manufactured items
    - If 0 manufactured → use header (add -001 suffix if missing)
    - If 1 manufactured → use that item's part number
    - If 2+ manufactured → return "MULTIPLE"
    
    Returns:
        Part number string, "MULTIPLE", or empty string
    """
    if not pl_text or not pl_text.strip():
        return ''
    
    pl_format = _detect_pl_format(pl_text)
    if pl_format == 'UNKNOWN':
        return ''
    
    header_pn = _extract_header_part_number(pl_text, pl_format)
    manufactured = _extract_manufactured_items_from_text(pl_text, pl_format)
    
    if len(manufactured) == 0:
        if not header_pn:
            # Ultimate fallback: search entire text for -XXX pattern
            pn_match = re.search(r'\b([A-Z0-9]{4,}-\d{3})\b', pl_text[:1500])
            if pn_match:
                candidate = pn_match.group(1)
                if candidate.lower() not in _NOT_PART_NUMBERS:
                    return candidate
            return ''
        # If header already has a -XXX suffix, return as-is
        if _has_suffix(header_pn):
            return header_pn
        # No manufactured items and no suffix → search for item that contains header
        all_items_in_pl = re.findall(
            r'^\s*\d{1,3}\s+(\S+)\s+\w{1,3}\s+',
            pl_text, re.MULTILINE
        )
        candidates = []
        header_norm = re.sub(r'[\s_\-]', '', header_pn).lower()
        for item_pn in all_items_in_pl:
            item_norm = re.sub(r'[\s_\-]', '', item_pn).lower()
            if header_norm in item_norm and re.search(r'-\d{3}$', item_pn):
                candidates.append(item_pn)

        if len(candidates) == 1:
            # Exactly one matching item → use it
            logger.info(f"PL: header '{header_pn}' → single matching item '{candidates[0]}'")
            return candidates[0]
        elif len(candidates) > 1:
            # Multiple → MULTIPLE (don't guess)
            logger.info(f"PL: header '{header_pn}' → multiple candidates: {candidates}")
            return 'MULTIPLE'
        else:
            # No matching items → return header as-is (don't add -001)
            logger.info(f"PL: header '{header_pn}' → no matching items, returning as-is")
            return header_pn
    
    if len(manufactured) == 1:
        pn = manufactured[0]
        # Final validation: reject if it's a common word, not a real part number
        if pn.lower() in _NOT_PART_NUMBERS:
            # Fall back to header
            if header_pn and re.search(r'-\d{3}$', header_pn):
                return header_pn
            return ''
        # Must end with -XXX
        if re.search(r'-\d{3}$', pn):
            return pn
        # Manufactured item doesn't match -XXX, fall back to header
        if header_pn and re.search(r'-\d{3}$', header_pn):
            return header_pn
        return ''
    
    # 2+ manufactured items
    return 'MULTIPLE'


# ────────────────────────────────────────────────────────────────────────
def _generate_pl_summary_english(part_number: str, pl_items: List[Dict], ocr_original: str = '') -> str:
    """
    Generate English summary of PL items for a specific part number.
    FILTERED: Shows ONLY specific processes and materials
    
    Processes included: expected, painting, testing, anodizing, passivation, masking
    Materials included: inserts (fasteners), colors
    
    Args:
        part_number: Part number to match (from drawing)
        pl_items: List of PL item dictionaries
        ocr_original: Original OCR part number before PL override (if applicable)
    
    Returns:
        English summary string format: "process - spec | process - spec | color | insert"
    """
    if not pl_items or not part_number:
        return ""
    
    # Whitelist of allowed processes
    ALLOWED_PROCESSES = {
        'expected', 'צפוי',
        'painting', 'צביעה', 'paint',
        'testing', 'בדיקה', 'ביקורת', 'inspection', 'check',
        'packing', 'אריזה', 'pack',
        'passivation', 'פסיבציה', 'passivate',
        'anodizing', 'אנודייז', 'anodize',
        'masking', 'מיסוך', 'mask',
        'marking', 'סימון', 'mark',
        'laser marking', 'סימון בלייזר', 'laser',
        'coating', 'ציפוי',
        'machining', 'עיבוד', 'machine'
    }
    
    # Normalize the part number for comparison
    part_num_norm = _normalize_item_number(str(part_number))
    ocr_orig_norm = _normalize_item_number(str(ocr_original)) if ocr_original else ''
    
    matched_items = []
    for pl_item in pl_items:
        associated = pl_item.get('associated_item', '')
        associated_norm = _normalize_item_number(str(associated))
        
        matched = (associated_norm == part_num_norm or
            associated_norm in part_num_norm or
            part_num_norm in associated_norm)
        # Also match against original OCR PN (before PL override)
        if not matched and ocr_orig_norm:
            matched = (associated_norm == ocr_orig_norm or
                associated_norm in ocr_orig_norm or
                ocr_orig_norm in associated_norm)
        if matched:
            matched_items.append(pl_item)
    
    if not matched_items:
        return ""
    
    # Collect ALL process-spec pairs and materials (NO FILTERING)
    process_spec_pairs = []
    materials = set()
    
    for item in matched_items:
        # Get English processes (original from AI) - NO FILTERING, show all
        processes = item.get('processes', [])
        if isinstance(processes, str):
            processes = [processes] if processes else []
        
        specifications = item.get('specifications', [])
        if isinstance(specifications, str):
            specifications = [specifications] if specifications else []
        
        description = (item.get('description') or '').strip().lower()
        item_type = (item.get('item_type') or '').strip().lower()
        product_tree = (item.get('product_tree') or '').strip()
        
        # Pair ALL processes with specs (no filtering)
        if processes and specifications:
            for i, process in enumerate(processes):
                spec = specifications[i] if i < len(specifications) else (specifications[0] if specifications else "")
                if spec:
                    pair = (process.strip(), spec.strip())
                    if pair not in process_spec_pairs:
                        process_spec_pairs.append(pair)
                else:
                    pair = (process.strip(), "")
                    if pair not in process_spec_pairs:
                        process_spec_pairs.append(pair)
        elif processes:
            for process in processes:
                pair = (process.strip(), "")
                if pair not in process_spec_pairs:
                    process_spec_pairs.append(pair)
        
        # Extract materials (colors and inserts)
        if product_tree:
            tree_lower = product_tree.lower()
            if 'insert' in tree_lower or 'fastener' in tree_lower:
                materials.add(description) if description else None
            if 'color' in tree_lower:
                materials.add(description) if description else None
        
        if item_type:
            if 'color' in item_type or 'coil_insert' in item_type or 'fastener' in item_type:
                materials.add(description) if description else None
    
    # Build summary string
    summary_parts = []
    
    for process, spec in process_spec_pairs:
        if process and spec:
            summary_parts.append(f"{process} - {spec}")
        elif process:
            summary_parts.append(process)
        elif spec:
            summary_parts.append(spec)
    
    for material in sorted(materials):
        if material and material.strip():
            summary_parts.append(material)
    
    return " | ".join(summary_parts) if summary_parts else ""


# ────────────────────────────────────────────────────────────────────────
def _extract_structured_pl_fields_ai(pl_text: str, client: AzureOpenAI, header_pn: str = '') -> Tuple[dict, int, int]:
    """Extract structured PL fields using AI: processes, paint, hardware.
    
    Separate API call — does NOT replace existing extract_pl_data.
    
    Args:
        pl_text: Raw text from PL document
        client: Azure OpenAI client
        header_pn: Main part number from PL header (for self-reference filtering)
    
    Returns:
        Tuple of (fields_dict, input_tokens, output_tokens)
        fields_dict has keys: pl_processes, pl_paint, pl_hardware, pl_material, pl_summary_hebrew
    """
    empty = {'pl_processes': '', 'pl_paint': '', 'pl_hardware': '', 'pl_material': '', 'pl_summary_hebrew': ''}
    
    if not pl_text or not pl_text.strip():
        return empty, 0, 0
    
    prompt = load_prompt("11_extract_pl_fields") + "\n" + pl_text[:8000]

    try:
        import json
        
        stage6_model, stage6_max_tokens, stage6_temperature = _resolve_stage_call_config(
            STAGE_PL, 1000, 0.1
        )
        
        response = _chat_create_with_token_compat(
            client,
            model=stage6_model,
            messages=[
                {"role": "system", "content": "Extract structured data from Parts List documents. Return only valid JSON. Be concise."},
                {"role": "user", "content": prompt}
            ],
            temperature=stage6_temperature,
            max_tokens=stage6_max_tokens
        )
        
        tokens_in = response.usage.prompt_tokens
        tokens_out = response.usage.completion_tokens
        
        response_text = response.choices[0].message.content.strip()
        
        # Parse JSON
        json_start = response_text.find('{')
        json_end = response_text.rfind('}') + 1
        if json_start >= 0 and json_end > json_start:
            parsed = json.loads(response_text[json_start:json_end])
        else:
            parsed = json.loads(response_text)
        
        # ── PL Processes (English, short) ──
        proc_parts = []
        for p in parsed.get('processes', []):
            name = p.get('name', '')
            spec = p.get('spec', '')
            if name and spec:
                proc_parts.append(f"{name} — {spec}")
            elif name:
                proc_parts.append(name)

        # ── PL Paint (English, short) ──
        paint_parts = []
        for p in parsed.get('paint', []):
            spec = p.get('spec', '')
            color = p.get('color', '')
            entry = spec
            if color:
                entry += f" ({color})"
            if entry.strip():
                paint_parts.append(entry.strip())

        # ── PL Hardware (PN + qty) — with filtering and grouping ──
        header_pn_norm = re.sub(r'[\s_\-]', '', (header_pn or '')).lower()

        # Pre-filter: validate hardware items are real inserts (not sub-assembly parts)
        raw_hardware = parsed.get('hardware', [])
        validated_hw = validate_inserts_hardware(
            [{'cat_no': h.get('pn', ''), 'qty': h.get('qty', ''), 'description': h.get('desc', '') or h.get('pn', '')} for h in raw_hardware if isinstance(h, dict)],
            part_number=header_pn or 'PL',
        )
        # Rebuild hardware list with validated PNs only
        validated_pns = {item['cat_no'] for item in validated_hw}
        hardware_items = [h for h in raw_hardware if isinstance(h, dict) and h.get('pn', '').strip() in validated_pns]

        # Pass 1: classify each hardware item (primary vs alternative)
        items_raw = []   # [(pn, qty, is_alt)]

        for h in hardware_items:
            pn = h.get('pn', '').strip()
            qty = h.get('qty', '')
            seq = h.get('seq', '').strip()
            if not pn:
                continue

            pn_norm = re.sub(r'[\s_\-]', '', pn).lower()

            # Skip self-reference (main PN appearing as hardware)
            if header_pn_norm and pn_norm == header_pn_norm:
                logger.debug(f"PL Hardware: skipping self-reference '{pn}'")
                continue

            # GL_alt prefix — alternative
            if pn.upper().startswith('GL_ALT') or pn.upper().startswith('GL ALT'):
                alt_pn = re.sub(r'^GL[_ ]?ALT\s*', '', pn, flags=re.IGNORECASE).strip()
                if alt_pn:
                    items_raw.append((alt_pn, qty, True))
                continue

            # Seq-based: "Alt" / "Alternate" in seq column → alternative
            is_alt = bool(seq and re.match(r'(?i)^alt', seq))
            items_raw.append((pn, qty, is_alt))

        # Pass 2: group — alts attach to nearest primary (prefer preceding, fallback to next)
        primary_items = []   # [(pn, qty)]
        alt_map = {}         # primary_index → [(alt_pn, qty)]
        pending_alts = []    # alts seen before any primary
        current_primary_idx = -1

        for pn, qty, is_alt in items_raw:
            if not is_alt:
                primary_items.append((pn, qty))
                current_primary_idx = len(primary_items) - 1
                # Assign any pending alts (seen before first primary)
                if pending_alts:
                    alt_map.setdefault(current_primary_idx, []).extend(pending_alts)
                    pending_alts = []
            else:
                if current_primary_idx >= 0:
                    alt_map.setdefault(current_primary_idx, []).append((pn, qty))
                else:
                    pending_alts.append((pn, qty))

        # Leftover alts with no primary → treat as primary (fallback)
        for pn, qty in pending_alts:
            primary_items.append((pn, qty))

        # Pass 3: format grouped output (with price lookup)
        hw_parts = []
        hw_heb_items = []  # for summary (PN + qty + price for Stage 9 bom)

        def _format_hw_entry(pn: str, qty: str, suffix: str = '') -> str:
            """Format a hardware entry with optional price."""
            price_info = lookup_insert_price(pn)
            base = f"{pn}{suffix} ×{qty}" if qty else f"{pn}{suffix}"
            if price_info:
                price, currency = price_info
                base += f" ×{price}{currency}"
            return base

        for idx, (pn, qty) in enumerate(primary_items):
            alts = alt_map.get(idx, [])
            if alts:
                group_parts = [_format_hw_entry(pn, qty)]
                hw_heb_items.append(_format_hw_entry(pn, qty))
                for alt_pn, alt_qty in alts:
                    group_parts.append(_format_hw_entry(alt_pn, alt_qty, ' (חלופי)'))
                    hw_heb_items.append(_format_hw_entry(alt_pn, alt_qty, ' (חלופי)'))
                hw_parts.append(', '.join(group_parts))
            else:
                hw_parts.append(_format_hw_entry(pn, qty))
                hw_heb_items.append(_format_hw_entry(pn, qty))

        # ── PL Material (raw material from PARTS LIST) ──
        mat_parts = []
        mat_heb_items = []
        for m in parsed.get('material', []):
            pn = m.get('pn', '').strip()
            desc = m.get('desc', '').strip()
            spec = m.get('spec', '').strip()
            qty = m.get('qty', '')
            if not pn and not desc:
                continue
            entry_parts = []
            if pn:
                entry_parts.append(pn)
            if desc:
                entry_parts.append(desc)
            if spec:
                entry_parts.append(spec)
            if qty:
                entry_parts.append(f"×{qty}")
            mat_parts.append(' | '.join(entry_parts))
            # Hebrew summary: short desc or pn
            mat_heb_items.append(desc if desc else pn)

        # ── PL Summary Hebrew (combined with section headers) ──
        heb_sections = []
        
        # ציפוי — סוג + מפרט
        coating_items = []
        for p in parsed.get('processes', []):
            name_heb = p.get('name_heb', '') or p.get('name', '')
            spec = p.get('spec', '')
            if name_heb and spec:
                coating_items.append(f"{name_heb} {spec}")
            elif name_heb:
                coating_items.append(name_heb)
        if coating_items:
            heb_sections.append(f"ציפוי: {', '.join(coating_items)}")
        
        # צבעים — סוג + מפרט + גוון
        paint_items = []
        for p in parsed.get('paint', []):
            type_heb = p.get('type_heb', '')
            spec = p.get('spec', '')
            color = p.get('color', '')
            parts = []
            if type_heb:
                parts.append(type_heb)
            if spec:
                parts.append(spec)
            if color:
                parts.append(f"גוון {color}")
            if parts:
                paint_items.append(' '.join(parts))
        if paint_items:
            heb_sections.append(f"צבעים: {', '.join(paint_items)}")
        
        # קשיחים — filtered and grouped
        if hw_heb_items:
            heb_sections.append(f"קשיחים: {', '.join(hw_heb_items)}")

        # חומר גלם — from material category
        if mat_heb_items:
            heb_sections.append(f"חומר: {', '.join(mat_heb_items)}")
        
        heb_parts = heb_sections  # For the join below

        result = {
            'pl_processes': ' | '.join(proc_parts),
            'pl_paint': ' | '.join(paint_parts),
            'pl_hardware': ' | '.join(hw_parts),
            'pl_material': ' | '.join(mat_parts),
            'pl_summary_hebrew': ' | '.join(heb_parts),
        }
        
        return result, tokens_in, tokens_out
        
    except Exception as e:
        logger.error(f"WARNING: Structured PL extraction failed: {str(e)}")
        return empty, 0, 0


# ────────────────────────────────────────────────────────────────────────
def extract_pl_data(pl_pdf_path: str, client: AzureOpenAI, file_classifications: list = None) -> Tuple[list, int, int]:
    """
    Extract detailed information from Parts List (PL) PDF using Azure OpenAI.
    
    Args:
        pl_pdf_path: Path to the PL PDF file
        client: AzureOpenAI client instance
        file_classifications: List of file classification dicts to find associated_item
        
    Returns:
        Tuple of (items, input_tokens, output_tokens)
        - items: List of dicts with keys:
          - pl_filename: Source PDF filename
          - item_number: Part/assembly number
          - description: Item description
          - quantity: Item quantity (if available)
          - processes: List of processes (e.g., [machining, passivation, coating])
          - specifications: List of specs/standards (e.g., [AMS 2700, MIL-P-53022])
          - product_tree: Flat text showing hierarchy (e.g., "Parent: A1, Children: B1|B2|C1")
          - item_type: Type indicator (fastener, coil_insert, color, assembly, etc.)
        - input_tokens: Tokens used in request
        - output_tokens: Tokens used in response
    """
    try:
        import pdfplumber
    except ImportError:
        logger.warning("WARNING: pdfplumber not installed, skipping PL: " + pl_pdf_path)
        return []
    
    try:
        # Extract all text from PDF
        pl_text = ""
        try:
            with pdfplumber.open(pl_pdf_path) as pdf:
                for page in pdf.pages:
                    pl_text += page.extract_text() + "\n"
        except Exception as e:
            logger.error(f"ERROR reading PDF {pl_pdf_path}: {str(e)}")
            return []
        
        if not pl_text.strip():
            logger.warning(f"WARNING: No text extracted from PL: {pl_pdf_path}")
            return []
        
        # Determine main part number from PL (deterministic, no API cost)
        pl_main_part_number = _determine_pl_main_part_number(pl_text)
        pl_format = _detect_pl_format(pl_text)
        if pl_main_part_number:
            logger.info(f"PL main part number: {pl_main_part_number} (format: {pl_format})")
        
        # Prepare PL analysis prompt
        pl_analysis_prompt = load_prompt("12_analyze_pl_bom") + "\n" + pl_text[:8000]
        
        # Call Azure OpenAI
        stage6_model, stage6_max_tokens, stage6_temperature = _resolve_stage_call_config(
            STAGE_PL,
            2000,
            0.2,
        )
        response = _chat_create_with_token_compat(
            client,
            model=stage6_model,
            messages=[
                {"role": "system", "content": "You are an expert in analyzing technical Parts Lists and Bills of Materials. Extract and structure data in JSON format."},
                {"role": "user", "content": pl_analysis_prompt}
            ],
            temperature=stage6_temperature,
            max_tokens=stage6_max_tokens
        )
        
        # Track token usage
        usage = response.usage
        tokens_in = usage.prompt_tokens
        tokens_out = usage.completion_tokens
        
        response_text = response.choices[0].message.content.strip()
        
        # Parse JSON response
        try:
            # Try to extract JSON from response (in case there's extra text)
            import json
            import re as _re
            json_start = response_text.find('[')
            json_end = response_text.rfind(']') + 1
            if json_start >= 0 and json_end > json_start:
                json_str = response_text[json_start:json_end]
            else:
                json_str = response_text
            
            # Attempt 1: parse as-is
            try:
                items = json.loads(json_str)
            except json.JSONDecodeError:
                # Attempt 2: fix common LLM JSON issues
                fixed = json_str
                # Fix unescaped newlines inside strings
                fixed = _re.sub(r'(?<=": ")(.*?)(?=")', lambda m: m.group(0).replace('\n', ' '), fixed, flags=_re.DOTALL)
                # Fix trailing commas before ] or }
                fixed = _re.sub(r',\s*([}\]])', r'\1', fixed)
                # Fix missing commas between } and {
                fixed = _re.sub(r'}\s*{', '},{', fixed)
                try:
                    items = json.loads(fixed)
                except json.JSONDecodeError:
                    # Attempt 3: extract individual objects with regex
                    obj_pattern = _re.compile(r'\{[^{}]*\}', _re.DOTALL)
                    raw_objects = obj_pattern.findall(json_str)
                    items = []
                    for raw_obj in raw_objects:
                        clean_obj = raw_obj.replace('\n', ' ')
                        clean_obj = _re.sub(r',\s*}', '}', clean_obj)
                        try:
                            items.append(json.loads(clean_obj))
                        except json.JSONDecodeError:
                            continue
                    if not items:
                        raise
                    logger.warning(f"⚠️ Recovered {len(items)} items from malformed JSON (regex fallback)")
        except Exception as e:
            logger.error(f"ERROR parsing PL JSON response: {str(e)}")
            logger.info(f"Response snippet: {response_text[:200]}")
            return [], tokens_in, tokens_out
        
        # Add filename and associated_item to each item and clean up
        pl_filename = Path(pl_pdf_path).name
        
        # Find associated_item from file_classifications by matching filename
        associated_item = ""
        if file_classifications:
            for fc in file_classifications:
                if fc.get('original_filename', '') == pl_filename:
                    associated_item = fc.get('associated_item', '')
                    logger.info(f"Found associated_item '{associated_item}' for {pl_filename}")
                    break
        
        result_items = []
        
        # Translation dictionary for processes
        PROCESS_TRANSLATIONS = {
            'painting': 'צביעה',
            'paint': 'צביעה',
            'packing': 'אריזה',
            'pack': 'אריזה',
            'passivation': 'פסיבציה',
            'passivate': 'פסיבציה',
            'anodizing': 'אנודייז',
            'anodize': 'אנודייז',
            'testing': 'בדיקה',
            'inspection': 'בדיקה',
            'check': 'בדיקה',
            'masking': 'מיסוך',
            'mask': 'מיסוך',
            'marking': 'סימון',
            'mark': 'סימון',
            'laser marking': 'סימון בלייזר',
            'laser': 'סימון בלייזר',
            'coating': 'ציפוי',
            'machining': 'עיבוד',
            'machine': 'עיבוד',
            'expected': 'צפוי'
        }
        
        for item in items:
            if isinstance(item, dict):
                item['pl_filename'] = pl_filename
                item['associated_item'] = associated_item  # From file_classifications
                item['matched_item_name'] = item.get('item_number', '')  # Will be matched later
                item['pl_main_part_number'] = pl_main_part_number  # Main part number from PL header
                
                # Ensure lists are properly formatted
                if not isinstance(item.get('processes'), list):
                    item['processes'] = []
                if not isinstance(item.get('specifications'), list):
                    item['specifications'] = []
                
                # Translate processes to Hebrew
                processes_list = item.get('processes', [])
                processes_hebrew = []
                for process in processes_list:
                    process_lower = process.lower().strip()
                    # Try to find exact match or substring match
                    translated = None
                    for key, value in PROCESS_TRANSLATIONS.items():
                        if key in process_lower:
                            translated = value
                            break
                    processes_hebrew.append(translated if translated else process)
                
                item['processes_hebrew'] = processes_hebrew
                
                result_items.append(item)
        
        # ── Additional AI call: structured fields (processes/paint/hardware) ──
        try:
            structured, struct_tokens_in, struct_tokens_out = _extract_structured_pl_fields_ai(pl_text, client, header_pn=pl_main_part_number)
            tokens_in += struct_tokens_in
            tokens_out += struct_tokens_out
            
            # Attach structured fields to ALL items from this PL
            for item in result_items:
                item['pl_processes'] = structured.get('pl_processes', '')
                item['pl_paint'] = structured.get('pl_paint', '')
                item['pl_hardware'] = structured.get('pl_hardware', '')
                item['pl_material'] = structured.get('pl_material', '')
                item['pl_summary_hebrew'] = structured.get('pl_summary_hebrew', '')
            
            parts_found = []
            if structured.get('pl_processes'): parts_found.append('processes')
            if structured.get('pl_paint'): parts_found.append('paint')
            if structured.get('pl_hardware'): parts_found.append('hardware')
            if structured.get('pl_material'): parts_found.append('material')
            if parts_found:
                logger.info(f"Structured PL: {', '.join(parts_found)}")
        except Exception as struct_err:
            logger.error(f"WARNING: Structured PL fields error: {struct_err}")
            for item in result_items:
                item['pl_processes'] = ''
                item['pl_paint'] = ''
                item['pl_hardware'] = ''
                item['pl_material'] = ''
                item['pl_summary_hebrew'] = ''
        
        return result_items, tokens_in, tokens_out
        
    except Exception as e:
        logger.error(f"EXCEPTION in extract_pl_data: {str(e)}")
        return [], 0, 0
