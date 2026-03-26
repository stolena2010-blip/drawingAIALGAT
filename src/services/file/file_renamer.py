"""
File Renamer — Rename classified files with B2B prefixes
=========================================================

Extracted from customer_extractor_v3_dual.py (Step 7 of refactoring plan).

Functions:
  - rename_files_by_classification: Rename files based on type + associated item
"""

from pathlib import Path
from typing import Dict, List

from src.utils.logger import get_logger

logger = get_logger(__name__)

# Image extensions that should be reclassified as 3D_IMAGE
_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif', '.webp'}

# File-type → (prefix, type_suffix)
_TYPE_MAP: Dict[str, tuple] = {
    'DRAWING':  ('B2BDraw',  '_30'),
    '3D_MODEL': ('B2BModel', '_25'),
    '3D_IMAGE': ('B2BImg',   '_16'),
}
_DEFAULT_TYPE = ('B2BDoc', '_99')  # PARTS_LIST, OTHER, QUOTE, INVOICE, etc.


def rename_files_by_classification(file_classifications: List[Dict]) -> int:
    """
    Rename files in-place based on their classification and associated item.

    Updates each *fc* dict with ``file_path``, ``renamed_filename``, and
    optionally ``file_type`` (image-extension correction).

    Returns the number of files successfully renamed.
    """
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

        associated_item = (fc.get('associated_item') or '').strip()
        if not associated_item:
            continue

        file_type = fc.get('file_type', 'OTHER')

        # Fix: if file is actually an image, correct the classification
        ext_lower = source_path.suffix.lower()
        if ext_lower in _IMAGE_EXTENSIONS and file_type != '3D_IMAGE':
            logger.info(f"🔧 Correcting file type for {source_path.name}: {file_type} → 3D_IMAGE")
            file_type = '3D_IMAGE'
            fc['file_type'] = '3D_IMAGE'

        # Determine prefix and type suffix
        prefix, type_suffix = _TYPE_MAP.get(file_type, _DEFAULT_TYPE)

        # Build new filename: B2B{Type}_{original_name}{type_suffix}.{ext}
        original_name_no_ext = source_path.stem
        ext = source_path.suffix
        new_filename = f"{prefix}_{original_name_no_ext}{type_suffix}{ext}"
        new_path = source_path.parent / new_filename

        # Rename in source folder
        if source_path != new_path:
            try:
                source_path.rename(new_path)
                renamed_count += 1
                logger.info(f"✓ {source_path.name} → {new_filename}")
                # Update the file_path in classifications
                fc['file_path'] = new_path
            except Exception as e:
                logger.error(f"✗ Error renaming {source_path.name}: {e}")

        # Always store renamed filename for downstream mapping
        fc['renamed_filename'] = new_filename

    return renamed_count
