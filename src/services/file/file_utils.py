"""
File Utilities – metadata, PDF detection, drawing association, renaming, copying
================================================================================
Extracted from customer_extractor_v3_dual.py  (Phase 2.6)
"""

import json
import re
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Any, Tuple

import pdfplumber

from src.core.constants import (
    debug_print,
    WARN_FILE_SIZE_MB,
    MAX_FILE_SIZE_MB,
)
from src.services.extraction.filename_utils import _extract_item_number_from_filename
from src.utils.logger import get_logger

logger = get_logger(__name__)


# ── public API ──────────────────────────────────────────────────────────
__all__ = [
    "_get_file_metadata",
    "_detect_text_heavy_pdf",
    "_build_drawing_part_map",
    "_find_associated_drawing",
    "_rename_files_by_classification",
    "_create_metadata_json",
    "_copy_folder_to_tosend",
]


# ────────────────────────────────────────────────────────────────────────
def _get_file_metadata(file_path: Path) -> Dict[str, Any]:
    """Extract file metadata including display name, title, etc.
    
    Returns dict with:
        - file_name: actual filename
        - display_name: Windows display name (if available)
        - file_title: PDF/Office document title (if available)
        - file_size_mb: file size in MB
        - created_date: creation date
        - modified_date: modification date
        - is_large_file: True if file > WARN_FILE_SIZE_MB
        - is_too_large: True if file > MAX_FILE_SIZE_MB
    """
    metadata = {
        'file_name': file_path.name,
        'display_name': '',
        'file_title': '',
        'file_size_mb': 0,
        'created_date': '',
        'modified_date': '',
        'is_large_file': False,
        'is_too_large': False
    }
    
    try:
        # File size
        file_size = file_path.stat().st_size
        file_size_mb = round(file_size / (1024 * 1024), 2)
        metadata['file_size_mb'] = file_size_mb
        
        # Check if file is large or too large
        metadata['is_large_file'] = file_size_mb > WARN_FILE_SIZE_MB
        metadata['is_too_large'] = file_size_mb > MAX_FILE_SIZE_MB
        
        # Dates
        created = datetime.fromtimestamp(file_path.stat().st_ctime)
        modified = datetime.fromtimestamp(file_path.stat().st_mtime)
        metadata['created_date'] = created.strftime("%Y-%m-%d %H:%M:%S")
        metadata['modified_date'] = modified.strftime("%Y-%m-%d %H:%M:%S")
        
        # Try to get document title from PDF metadata
        ext = file_path.suffix.lower()
        if ext == '.pdf':
            try:
                with pdfplumber.open(file_path) as pdf:
                    if pdf.metadata and pdf.metadata.get('Title'):
                        metadata['file_title'] = pdf.metadata['Title']
                        # If title is different from filename, use it as display name
                        if metadata['file_title'] and metadata['file_title'] != file_path.stem:
                            metadata['display_name'] = metadata['file_title']
            except Exception as e:
                debug_print(f"[PDF_METADATA_TITLE_EXTRACTION] Error: {e}")
                pass
        
        # For Office files, try to get document properties
        elif ext in {'.docx', '.xlsx', '.pptx'}:
            try:
                # This requires python-docx or openpyxl
                if ext == '.xlsx':
                    from openpyxl import load_workbook
                    wb = load_workbook(file_path, read_only=True)
                    if wb.properties.title:
                        metadata['file_title'] = wb.properties.title
                        if metadata['file_title'] != file_path.stem:
                            metadata['display_name'] = metadata['file_title']
                    wb.close()
            except Exception as e:
                debug_print(f"[OFFICE_METADATA_TITLE_EXTRACTION] Error: {e}")
                pass
        
        # Windows alternate name (8.3 format) - not really display name
        # Real display name would require Windows API calls which are complex
        
    except Exception as e:
        pass
    
    return metadata


# ────────────────────────────────────────────────────────────────────────
def _detect_text_heavy_pdf(file_path: Path) -> Tuple[bool, str]:
    """
    Heuristic guardrail: detect instruction-like PDFs that are mostly text/table
    with almost no geometric elements, to avoid misclassifying them as drawings.
    """
    try:
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return False, ""

            page = pdf.pages[0]
            text = page.extract_text() or ""
            if not text:
                return False, ""

            word_count = len(text.split())
            line_count = text.count("\n") + 1
            vector_elements = (
                len(getattr(page, "lines", []))
                + len(getattr(page, "curves", []))
                + len(getattr(page, "rects", []))
            )
            image_elements = len(getattr(page, "images", []))
            avg_chars_per_line = len(text) / max(line_count, 1)

            # Text-heavy signal: many words/lines, long lines, almost no vectors/images
            text_to_vector_ratio = len(text) / max(vector_elements + image_elements * 5, 1)

            if (
                word_count >= 250
                and line_count >= 40
                and avg_chars_per_line >= 30
                and vector_elements < 40
                and image_elements == 0
                and text_to_vector_ratio > 40
            ):
                reason = (
                    f"Text-heavy first page ({word_count} words, {line_count} lines, "
                    f"{vector_elements} vector elements, no images)"
                )
                return True, reason

        return False, ""
    except Exception as e:
        # If detection fails, fallback to normal classification flow
        logger.debug(f"Text-heavy detection skipped: {e}")
        return False, ""


# ────────────────────────────────────────────────────────────────────────
def _build_drawing_part_map(file_classifications: List[Dict], drawing_results: Optional[List[Dict]] = None) -> Dict[str, str]:
    """
    Build mapping of item_number (extracted from filename) -> part_number.
    All files with same item_number belong to same part.
    """
    mapping = {}
    
    if not drawing_results:
        return mapping
    
    for fc in file_classifications or []:
        if fc.get('file_type') != 'DRAWING':
            continue
        
        # Find the drawing result for this file
        part_number = None
        file_name = fc['file_path'].name
        for dr in drawing_results:
            if dr.get('file_name') == file_name:
                part_number = dr.get('part_number')
                break
        
        if part_number:
            # Extract item number from filename
            item_number = _extract_item_number_from_filename(file_name)
            if item_number:
                mapping[item_number] = part_number
                logger.debug(f"Drawing map: {item_number} → {part_number} (from {file_name})")
            
            # Also map by filename (for pre-Stage6 PL matching that uses filename as key)
            mapping[file_name] = part_number
    
    return mapping


# ────────────────────────────────────────────────────────────────────────
def _find_associated_drawing(file_path: Path, file_type: str, drawing_map: Dict[str, Any]) -> str:
    """
    Find associated part number for a non-drawing file.
    
    Two checks for PLs:
      1. Filename similarity (prefix match)
      2. OCR part number comparison (PL filename vs drawing part number)
    
    drawing_map values can be:
      - str: part_number directly
      - dict: {'part_number': ..., 'drawing_number': ..., 'revision': ...}
    """
    if not file_path or file_type == 'DRAWING' or not drawing_map:
        return ""
    
    # Helper: extract part_number from map value (handles both str and dict)
    def _get_pn(val):
        if isinstance(val, dict):
            return val.get('part_number', '')
        return str(val) if val else ''
    
    # Extract item number from current file
    current_item_number = _extract_item_number_from_filename(file_path.name)
    if not current_item_number:
        return ""
    
    # ── Check 1: Exact item_number match ──
    if current_item_number in drawing_map:
        return _get_pn(drawing_map[current_item_number])
    
    # ── For PLs: also extract the part number hint from PL filename ──
    is_pl = file_type == 'PARTS_LIST' or file_path.name.upper().startswith('PL')
    pl_hint = ""
    if is_pl:
        pl_stem = file_path.stem.upper()
        pl_clean = re.sub(r'^PL[_\-]?', '', pl_stem)
        pl_clean = re.sub(r'[_\-][A-Z\-]{1,3}$', '', pl_clean)
        pl_clean = re.sub(r'[_\-]KBM$', '', pl_clean, flags=re.IGNORECASE)
        pl_hint = re.sub(r'[^a-z0-9]', '', pl_clean.lower())
    
    # ── Score all candidates ──
    current_clean = re.sub(r'[^a-z0-9]', '', current_item_number.lower())
    best_match = ""
    best_score = 0
    
    for map_key, map_val in drawing_map.items():
        pn = _get_pn(map_val)
        if not pn:
            continue
        
        key_clean = re.sub(r'[^a-z0-9]', '', str(map_key).lower())
        pn_clean = re.sub(r'[^a-z0-9]', '', pn.lower())
        
        if not key_clean or len(key_clean) < 4:
            continue
        
        score = 0
        
        # ── Check A: filename item_number overlap ──
        # How many characters from the start match?
        common = 0
        for a, b in zip(current_clean, key_clean):
            if a == b:
                common += 1
            else:
                break
        
        if common >= 5:
            score += common * 10
        
        # ── Check B (PL only): PL hint vs drawing part_number ──
        if is_pl and pl_hint and pn_clean:
            # Remove trailing 3-digit dash suffix (e.g. "001")
            pn_base = re.sub(r'\d{3}$', '', pn_clean)
            pl_base = re.sub(r'\d{3}$', '', pl_hint)
            
            if pn_base and pl_base:
                # Check containment
                if pn_base in pl_base or pl_base in pn_base:
                    score += 100  # Strong: PL filename contains drawing PN
                else:
                    # Prefix overlap
                    pn_common = 0
                    for a, b in zip(pn_base, pl_base):
                        if a == b:
                            pn_common += 1
                        else:
                            break
                    pn_overlap = pn_common / max(len(pn_base), len(pl_base))
                    if pn_overlap >= 0.7:
                        score += 60
                    elif pn_overlap >= 0.5:
                        score += 30
        
        if score > best_score:
            best_score = score
            best_match = pn
    
    # Minimum threshold
    if best_score >= 50:
        return best_match
    
    return ""


# ────────────────────────────────────────────────────────────────────────
def _rename_files_by_classification(target_folder: Path, file_classifications: List[Dict], file_path_map: Optional[Dict[str, str]] = None) -> Dict[str, str]:
    """
    Rename files in target folder based on classification and associated item.
    Pattern: B2B{Type}_{item_number}_{index}.{ext}
    Only renames files that have associated_item set.
    
    Returns mapping of old_filename -> new_filename
    """
    rename_map = {}
    if not file_classifications:
        return rename_map
    
    logger.info(f"Starting rename - total {len(file_classifications)} files")
    
    # Build counter for files with same associated_item
    item_counters = {}  # {associated_item: counter}
    
    for fc in file_classifications:
        file_path = fc['file_path']
        if not file_path.exists():
            logger.warning(f"⚠️ File does not exist: {file_path}")
            continue
        
        associated_item = fc.get('associated_item', '')
        if not associated_item:
            continue  # Skip files without associated item
        
        file_type = fc.get('file_type', 'OTHER')
        
        # תיקון חזק: אם קובץ הוא תמונה בפועל (jpg, png וכו'), תיקון הסיווג ל-3D_IMAGE
        ext_lower = file_path.suffix.lower()
        if ext_lower in {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif', '.webp'}:
            if file_type != '3D_IMAGE':
                file_type = '3D_IMAGE'
        
        # Determine prefix based on type
        if file_type == 'DRAWING':
            prefix = 'B2BDraw'
        elif file_type == '3D_MODEL':
            prefix = 'B2BModel'
        elif file_type == '3D_IMAGE':
            prefix = 'B2BImg'
        else:  # PARTS_LIST, OTHER, QUOTE, INVOICE, PURCHASE_ORDER
            prefix = 'B2BDoc'
        
        # Get counter for this item
        if associated_item not in item_counters:
            item_counters[associated_item] = 0
        item_counters[associated_item] += 1
        counter = item_counters[associated_item]
        
        # Build new filename
        ext = file_path.suffix.lower()
        new_filename = f"{prefix}_{associated_item}_{counter:02d}{ext}"
        
        # Rename file
        new_path = target_folder / new_filename
        try:
            if file_path != new_path:
                file_path.rename(new_path)
                fc['file_path'] = new_path  # Update path in classifications
                rename_map[file_path.name] = new_filename
                logger.info(f"✓ {file_path.name} → {new_filename}")
        except Exception as e:
            logger.error(f"Failed to rename {file_path.name}: {e}")
    
    if rename_map:
        logger.info(f"Renamed {len(rename_map)} files")
    else:
        logger.info("No files renamed (no associated_item)")
    
    return rename_map


# ────────────────────────────────────────────────────────────────────────
def _create_metadata_json(file_classifications: List[Dict], output_folder: Path, filename: str = "metadata.json") -> Path:
    """
    Create a JSON file with metadata (renamed_filename and display_name) for all files with associated_item.
    
    Args:
        file_classifications: List of file classifications
        output_folder: Folder where to save the JSON file
        filename: Name of the JSON file
        
    Returns:
        Path to created JSON file
    """
    metadata = {"files": []}
    
    for fc in file_classifications:
        associated_item = (fc.get('associated_item') or '').strip()
        if not associated_item:
            continue
        
        # Get renamed_filename (with B2B codes)
        file_name = fc.get('renamed_filename', '')
        if not file_name:
            # Fallback to original_filename if renamed_filename not available
            file_name = fc.get('original_filename', '')
        
        display_name = (fc.get('display_name') or '').strip()
        
        if file_name and display_name:
            metadata["files"].append({
                "file_name": file_name,
                "display_name": display_name
            })
    
    if metadata["files"]:
        output_file = output_folder / filename
        try:
            # Use cp1255 (Windows-1255) encoding to match B2B text files
            with open(output_file, 'w', encoding='cp1255', errors='replace') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)
            logger.info(f"✓ Created metadata: {filename} ({len(metadata['files'])} files)")
            return output_file
        except Exception as e:
            logger.error(f"Failed to create metadata: {e}")
            return None
    else:
        logger.warning(f"⚠ No files with associated_item - metadata not created")
        return None


# ────────────────────────────────────────────────────────────────────────
def _create_filtered_metadata_json(
    file_classifications: List[Dict], 
    results: List[Dict],
    output_folder: Path, 
    confidence_level: str = "HIGH",
    filename: str = "metadata.json"
) -> Path:
    """
    Create a filtered metadata.json — only files whose associated_item 
    has confidence matching the B2B filter.
    
    Args:
        file_classifications: List of file classifications
        results: List of drawing analysis results (with confidence_level)
        output_folder: Where to save the JSON file
        confidence_level: "HIGH" (HIGH+FULL), "MEDIUM" (MEDIUM+HIGH+FULL), "LOW" (all)
        filename: Output filename
    """
    import re
    
    # Build set of part numbers that pass the confidence filter
    allowed_confidence = set()
    if confidence_level == "LOW":
        allowed_confidence = {'LOW', 'MEDIUM', 'HIGH', 'FULL'}
    elif confidence_level == "MEDIUM":
        allowed_confidence = {'MEDIUM', 'HIGH', 'FULL'}
    else:  # HIGH (default)
        allowed_confidence = {'HIGH', 'FULL'}
    
    # Get part numbers that pass the filter
    allowed_parts = set()
    for result in (results or []):
        conf = str(result.get('confidence_level', '')).strip().upper()
        pn = str(result.get('part_number', '')).strip()
        if conf in allowed_confidence and pn:
            # Normalize for comparison
            allowed_parts.add(re.sub(r'[^A-Za-z0-9]', '', pn).upper())
    
    metadata = {"files": []}
    
    for fc in file_classifications:
        associated_item = (fc.get('associated_item') or '').strip()
        if not associated_item:
            continue
        
        file_name = fc.get('renamed_filename', '') or fc.get('original_filename', '')
        display_name = (fc.get('display_name') or '').strip()
        
        if not file_name or not display_name:
            continue
        
        # Check if this file's associated_item is in the allowed set
        assoc_norm = re.sub(r'[^A-Za-z0-9]', '', associated_item).upper()
        
        if confidence_level == "LOW" or assoc_norm in allowed_parts:
            metadata["files"].append({
                "file_name": file_name,
                "display_name": display_name
            })
        else:
            # Also check if any allowed part is a substring (for partial matches)
            matched = any(
                assoc_norm in ap or ap in assoc_norm 
                for ap in allowed_parts if len(ap) > 4
            )
            if matched:
                metadata["files"].append({
                    "file_name": file_name,
                    "display_name": display_name
                })
    
    if metadata["files"]:
        output_file = output_folder / filename
        try:
            with open(output_file, 'w', encoding='cp1255', errors='replace') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)
            logger.info(f"✓ Created filtered metadata: {filename} ({len(metadata['files'])} files, confidence={confidence_level})")
            return output_file
        except Exception as e:
            logger.error(f"Failed to create filtered metadata: {e}")
            return None
    else:
        logger.warning(f"⚠ No files passed confidence filter for {filename}")
        return None


# ────────────────────────────────────────────────────────────────────────
def _copy_folder_to_tosend(source_folder: Path, tosend_folder: Path, file_classifications: Optional[List[Dict]] = None, confidence_level: str = "FULL", results: Optional[List[Dict]] = None) -> bool:
    """
    Rename files in source folder based on classification, then copy entire folder to TO_SEND.
    Also handles B2B file selection based on confidence level.
    
    Args:
        source_folder: Source folder to process and copy
        tosend_folder: Target TO_SEND folder
        file_classifications: Optional list of file classifications for renaming
        confidence_level: B2B file confidence filter - "FULL" (only FULL), "HIGH" (HIGH+FULL), "MEDIUM" (MEDIUM+HIGH+FULL)
        results: Optional list of drawing analysis results (for filtered metadata)
    
    Returns:
        True if successful
    """
    try:
        if not tosend_folder.exists():
            tosend_folder.mkdir(parents=True, exist_ok=True)
        
        # Phase 1: Rename files in source folder based on classification
        if file_classifications:
            logger.info(f"[RENAME] Renaming files in source folder...")
            renamed_count = 0
            for fc in file_classifications:
                source_path = fc.get('file_path')
                if not source_path:
                    continue
                    
                # Convert to Path if string
                if isinstance(source_path, str):
                    source_path = Path(source_path)
                
                if not source_path.exists():
                    continue
                
                associated_item = fc.get('associated_item', '')
                file_type = fc.get('file_type', 'OTHER')
                
                # תיקון חזק: אם קובץ הוא תמונה בפועל (jpg, png וכו'), תיקון הסיווג ל-3D_IMAGE
                ext_lower = source_path.suffix.lower()
                if ext_lower in {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif', '.webp'}:
                    if file_type != '3D_IMAGE':
                        logger.info(f"🔧 Correcting file type for {source_path.name}: {file_type} → 3D_IMAGE")
                        file_type = '3D_IMAGE'
                        fc['file_type'] = '3D_IMAGE'
                
                # תיקון: תמונות צריכות להיות מעובדות אפילו ללא associated_item
                if not associated_item and file_type != '3D_IMAGE':
                    continue  # Skip non-image files without associated_item
                
                # Determine prefix and type suffix based on type
                if file_type == 'DRAWING':
                    prefix = 'B2BDraw'
                    type_suffix = '_30'
                elif file_type == '3D_MODEL':
                    prefix = 'B2BModel'
                    type_suffix = '_25'
                elif file_type == '3D_IMAGE':
                    prefix = 'B2BImg'
                    type_suffix = '_16'
                else:  # PARTS_LIST, OTHER, QUOTE, INVOICE, PURCHASE_ORDER
                    prefix = 'B2BDoc'
                    type_suffix = '_99'
                
                # Build new filename
                original_name_no_ext = source_path.stem
                ext = source_path.suffix
                
                # Check if already renamed (has B2B prefix) - if so, skip renaming
                if source_path.name.startswith(('B2BDraw_', 'B2BModel_', 'B2BImg_', 'B2BDoc_')):
                    logger.debug(f"⚠ Already renamed: {source_path.name} - skipping")
                    # Still update fc with the current filename
                    fc['renamed_filename'] = source_path.name
                    continue
                
                new_filename = f"{prefix}_{original_name_no_ext}{type_suffix}{ext}"
                new_path = source_folder / new_filename
                
                # Rename in source folder
                if source_path != new_path:
                    try:
                        source_path.rename(new_path)
                        fc['file_path'] = new_path  # update path after rename
                        fc['renamed_filename'] = new_filename
                        renamed_count += 1
                        logger.info(f"✓ {source_path.name} → {new_filename}")
                    except Exception as e:
                        logger.error(f"Error renaming {source_path.name}: {e}")
            
            logger.info(f"[RESULT] Renamed {renamed_count} files")
            
            # Create ALL_METADATA.json (unfiltered — all files, for RERUN)
            logger.info(f"[METADATA] Creating ALL_METADATA.json (all files)...")
            _create_metadata_json(file_classifications, source_folder, "ALL_METADATA.json")
            
            # Create metadata.json (filtered by confidence — matches B2B)
            if results and confidence_level != "LOW":
                logger.info(f"[METADATA] Creating metadata.json (filtered, confidence={confidence_level})...")
                _create_filtered_metadata_json(
                    file_classifications, results, source_folder, 
                    confidence_level, "metadata.json"
                )
            else:
                # LOW = all files, or no results available → same as ALL
                logger.info(f"[METADATA] Creating metadata.json (all files)...")
                _create_metadata_json(file_classifications, source_folder, "metadata.json")
        
        # Phase 1.5: Handle B2B file variants based on confidence level
        logger.info(f"[B2B] Processing B2B file variants (confidence: {confidence_level})...")
        b2b_files = list(source_folder.glob("B2B*-*.txt"))
        if b2b_files:
            # Determine which variant to keep as active
            if confidence_level == "LOW":
                target_variant = "B2B"
            elif confidence_level == "HIGH":
                target_variant = "B2BH"
            else:  # MEDIUM
                target_variant = "B2BM"
            
            # Step 0: Save the ALL variant (unfiltered B2B) as ALL_B2B for RERUN
            # Only needed if confidence is not LOW (LOW = already all rows)
            if confidence_level != "LOW":
                for bf in b2b_files:
                    prefix = bf.stem.split('-')[0]
                    if prefix == "B2B":  # This is the ALL variant (unfiltered)
                        all_name_parts = bf.stem.split('-')
                        all_new_name = f"ALL_B2B-{'-'.join(all_name_parts[1:])}.txt"
                        all_new_path = source_folder / all_new_name
                        try:
                            import shutil as _sh
                            _sh.copy2(bf, all_new_path)
                            logger.info(f"✓ Saved ALL variant for RERUN: {all_new_name}")
                        except Exception as e:
                            logger.error(f"Error saving ALL variant: {e}")
            
            # Step 1: Delete unwanted variant files (keep target + ALL_B2B)
            for bf in b2b_files:
                prefix = bf.stem.split('-')[0]  # e.g., "B2BH" or "B2B"
                if prefix != target_variant and not bf.name.startswith("ALL_B2B"):
                    try:
                        bf.unlink()
                        logger.debug(f"Removed {prefix}: {bf.name}")
                    except Exception as e:
                        logger.error(f"Error removing {bf.name}: {e}")
            
            # Step 2: Rename target variant to standard "B2B-0_" name (if needed)
            remaining_files = list(source_folder.glob("B2B[HM]-*.txt"))
            for bf in remaining_files:
                prefix = bf.stem.split('-')[0]
                if prefix == target_variant:
                    name_parts = bf.stem.split('-')
                    new_name = f"B2B-{'-'.join(name_parts[1:])}.txt"
                    new_path = source_folder / new_name
                    try:
                        bf.rename(new_path)
                        logger.info(f"✓ Selected {target_variant} → renamed to {new_name}")
                    except Exception as e:
                        logger.error(f"Error renaming {bf.name}: {e}")
        
        # Phase 2: Copy entire folder to TO_SEND
        dest_name = f"{source_folder.name}_TO_SEND"
        dest_path = tosend_folder / dest_name
        
        # If destination exists, remove it first
        if dest_path.exists():
            shutil.rmtree(dest_path)
        
        # Copy the entire folder (excluding ZIP files)
        def ignore_zips(directory, contents) -> list:
            return [f for f in contents if f.lower().endswith('.zip')]
        
        shutil.copytree(source_folder, dest_path, ignore=ignore_zips)
        
        # Count files (excluding classification files created during processing)
        skip_patterns = ('file_classification_', 'drawing_results_', 'SUMMARY_')
        file_count = sum(1 for f in dest_path.iterdir() if f.is_file() and not f.name.startswith(skip_patterns))
        
        logger.info(f"Copied to TO_SEND: {dest_name} ({file_count} files)")
        return True
    
    except Exception as e:
        logger.error(f"Failed to copy to TO_SEND: {e}")
        import traceback
        traceback.print_exc()
        return False
