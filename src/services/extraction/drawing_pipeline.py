"""
Drawing Data Extraction Pipeline — Stages 0-4
==============================================

Extracted from customer_extractor_v3_dual.py (Step 4 of refactoring plan).

Runs the full drawing analysis pipeline:
  - File validation & image extraction
  - OCR (Tesseract)
  - Customer detection (Rafael / IAI / Standard)
  - Stage 0.5: Layout identification
  - Stage 1: Basic info (title block)
  - Stage 2: Processes
  - Stage 3: NOTES
  - Stage 4: Geometric area
  - Post-processing & P.N. voting
"""

import os
import re
import base64
import json
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
from pathlib import Path
from typing import Optional, Tuple, Dict, Any

import pdfplumber
from openai import AzureOpenAI

from src.services.image.processing import (
    _downsample_high_res_image,
    _enhance_contrast_for_title_block,
    _extract_image_smart,
    _assess_image_quality,
    _apply_rotation_angle,
    _fix_image_rotation,
)
from src.services.ai.vision_api import (
    _call_vision_api_with_retry,
    _calculate_stage_cost,
)
from src.services.extraction import ocr_engine as _ocr_mod
from src.services.extraction.ocr_engine import (
    MultiOCREngine,
    extract_stage1_with_retry,
)
from src.services.extraction.stages_generic import (
    identify_drawing_layout,
    extract_basic_info,
    extract_processes_info,
    validate_notes_before_stage5,
    extract_notes_text,
    calculate_geometric_area,
)
from src.services.extraction.stages_rafael import (
    identify_drawing_layout_rafael,
    extract_basic_info_rafael,
    extract_processes_info_rafael,
    extract_processes_from_notes,
    extract_notes_text_rafael,
    extract_area_info_rafael,
)
from src.services.extraction.stages_iai import (
    identify_drawing_layout_iai,
    extract_basic_info_iai,
    extract_processes_info_iai,
    extract_notes_text_iai,
    extract_area_info_iai,
)
from src.services.extraction.pn_voting import (
    extract_pn_dn_from_text,
    vote_best_pn,
)
from src.services.extraction.post_processing import (
    post_process_summary_from_notes,
)

from src.utils.logger import get_logger

logger = get_logger(__name__)

# ── Constants ─────────────────────────────────────────────────────────
DRAWING_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}
MAX_IMAGE_DIMENSION = 4096
WARN_IMAGE_DIMENSION = 3000

STAGE_CLASSIFICATION = 0
STAGE_HARD_TIMEOUT = int(os.getenv("STAGE_HARD_TIMEOUT_SECONDS", "300"))


# ── Helpers ───────────────────────────────────────────────────────────

def _run_with_timeout(func, args=(), kwargs=None, timeout=STAGE_HARD_TIMEOUT, stage_name=""):
    """
    הרצת פונקציה עם hard timeout.
    אם הפונקציה תוקעת מעל timeout שניות — מחזיר None.
    """
    if kwargs is None:
        kwargs = {}
    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(func, *args, **kwargs)
        try:
            result = future.result(timeout=timeout)
            return result
        except FuturesTimeoutError:
            logger.error(
                f"\U0001f534 HARD TIMEOUT: {stage_name} exceeded {timeout}s ({timeout // 60} min) — skipping"
            )
            return None
        except Exception as e:
            logger.error(f"\U0001f534 {stage_name} failed: {e}")
            return None


# ── Main Pipeline ─────────────────────────────────────────────────────

def extract_drawing_data(
    file_path: str,
    client: AzureOpenAI,
    ocr_engine: MultiOCREngine,
    selected_stages: Dict[int, bool],
    enable_retry: bool = False,
    stage1_skip_retry_resolution_px: int = 8000,
) -> Tuple[Optional[Dict], Optional[Tuple[int, int]], Optional[Dict]]:
    """
    Main pipeline: extract all data from a single drawing file.

    Runs Stages 0-4, post-processing, and P.N. voting.

    Returns
    -------
    (result_data, (total_input_tokens, total_output_tokens), context)
        *result_data* — merged dict from all stages.
        *context* — dict with ``pdfplumber_text``, ``combined_text_stage1``,
                     ``is_rafael``, ``is_iai``, ``num_pages``.
        Returns ``(None, None, None)`` when the file should be skipped.
    """

    # ── timeout bookkeeping ──────────────────────────────────────
    try:
        file_timeout_seconds = int(os.getenv("FILE_TIMEOUT_SECONDS", "600") or 600)
    except Exception:
        file_timeout_seconds = 600
    extraction_start_time = time.time()
    timeout_triggered = False
    timeout_elapsed = 0.0

    def _check_timeout() -> bool:
        nonlocal timeout_triggered, timeout_elapsed
        if file_timeout_seconds <= 0:
            return False
        timeout_elapsed = time.time() - extraction_start_time
        if timeout_elapsed > file_timeout_seconds:
            timeout_triggered = True
            logger.warning(
                f"⚠️ TIMEOUT: {Path(file_path).name} took {timeout_elapsed:.0f}s "
                f"(limit: {file_timeout_seconds}s) — skipping remaining stages"
            )
            return True
        return False

    # ── file validation ──────────────────────────────────────────
    ext = Path(file_path).suffix.lower()

    if ext not in DRAWING_EXTS:
        logger.info(">> Skipped (not a supported drawing file): " + ext)
        return None, None, None

    # Quick check: skip non-drawing images (too small / odd aspect ratio)
    if ext in {".png", ".jpg", ".jpeg"}:
        try:
            from PIL import Image
            with Image.open(file_path) as img:
                width, height = img.size
                if width < 800 or height < 600:
                    logger.info(f">> Skipped (image too small for drawing: {width}x{height})")
                    return None, None, None
                aspect_ratio = max(width, height) / min(width, height)
                if aspect_ratio > 3:
                    logger.info(f">> Skipped (unusual aspect ratio: {width}x{height})")
                    return None, None, None
        except Exception as e:
            logger.info(f"Could not check image dimensions: {e}")

    # ── quality metadata ─────────────────────────────────────────
    quality_metadata: Dict[str, Any] = {
        'image_resolution': '',
        'quality_issues': []
    }

    # ── Step 1: Extract text and image ───────────────────────────
    text_snippet = ""
    img_bytes = None
    title_block_bytes = None
    additional_regions: Dict = {}
    num_pages = 1
    used_high_res = False
    was_downsampled = False
    orig_size = (0, 0)
    new_size = (0, 0)
    _pdfplumber_text = ""
    additional_pages_b64: list = []
    rotation_angle = 0  # declared early so image-file branch can reference it

    if ext == ".pdf":
        try:
            with pdfplumber.open(file_path) as pdf:
                text_snippet = "\n".join([(p.extract_text() or "") for p in pdf.pages[:3]])
                text_snippet = text_snippet[:5000]
                if text_snippet.strip():
                    logger.info(f"pdfplumber: {len(text_snippet)} chars")
                    _pdfplumber_text = text_snippet
        except Exception as e:
            logger.error(f"pdfplumber failed: {e}")

        try:
            import fitz
            with fitz.open(file_path) as doc:
                num_pages = len(doc)
                logger.info(f"PDF pages: {num_pages}")

            # Smart extraction with high-res support
            overview_bytes, tb_bytes, additional_regions, used_high_res = _extract_image_smart(file_path, page_num=1, dpi=200)

            # Multi-page: render additional pages for Stage 2/3
            if num_pages > 1:
                import fitz as fitz_mp
                try:
                    with fitz_mp.open(file_path) as doc_mp:
                        max_extra_pages = min(num_pages, 4)
                        for pg_idx in range(1, max_extra_pages):
                            page_mp = doc_mp[pg_idx]
                            pix_mp = page_mp.get_pixmap(dpi=150)
                            pg_bytes = pix_mp.tobytes("png")
                            pg_b64 = base64.b64encode(pg_bytes).decode('utf-8')
                            additional_pages_b64.append(pg_b64)
                    if additional_pages_b64:
                        logger.info(f"📄 Multi-page: rendered {len(additional_pages_b64)} additional page(s) for Stage 2/3")
                except Exception as e:
                    logger.warning(f"Multi-page rendering failed: {e}")
                    additional_pages_b64 = []

            # FIX ROTATION
            try:
                overview_bytes, was_rotated, rotation_info = _fix_image_rotation(overview_bytes, file_path)
                if was_rotated:
                    logger.info(f"📐 {rotation_info}")
                    quality_metadata['quality_issues'].append(rotation_info)
                    angle_match = re.search(r'(-?\d+)°', rotation_info)
                    if angle_match:
                        rotation_angle = int(angle_match.group(1))
                else:
                    logger.info(f"{rotation_info}")

                if tb_bytes and rotation_angle != 0:
                    tb_bytes = _apply_rotation_angle(tb_bytes, rotation_angle)

                if additional_regions and rotation_angle != 0:
                    for key in additional_regions:
                        if additional_regions[key] is not None:
                            additional_regions[key] = _apply_rotation_angle(additional_regions[key], rotation_angle)
            except Exception as e:
                logger.warning(f"⚠ Warning in PDF rotation fix: {e}")

            if used_high_res:
                logger.info(f"Using high-res region extraction")
                img_bytes = overview_bytes
                title_block_bytes = tb_bytes
                additional_regions = additional_regions

                import cv2
                import numpy as np
                nparr = np.frombuffer(overview_bytes, np.uint8)
                img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                if img_cv is not None:
                    orig_size = (img_cv.shape[1], img_cv.shape[0])
                    quality_metadata['image_resolution'] = f"{orig_size[0]}x{orig_size[1]}"
                    quality_metadata['quality_issues'].append("High-res: region extraction used")
            else:
                img_bytes = overview_bytes
                img_bytes, was_downsampled, orig_size, new_size = _downsample_high_res_image(img_bytes, MAX_IMAGE_DIMENSION)
                quality_metadata['image_resolution'] = f"{orig_size[0]}x{orig_size[1]}"

                if was_downsampled:
                    logger.info(f"Downsampled: {orig_size[0]}x{orig_size[1]}  {new_size[0]}x{new_size[1]}")
                    quality_metadata['quality_issues'].append("רזולוציה גבוהה - הופחתה אוטומטית")
                elif max(orig_size) > WARN_IMAGE_DIMENSION:
                    logger.info(f"High resolution: {orig_size[0]}x{orig_size[1]}")
                    quality_metadata['quality_issues'].append("רזולוציה גבוהה")

            # Assess image quality
            quality_assessment = _assess_image_quality(img_bytes)
            if quality_assessment['quality_issues']:
                quality_metadata['quality_issues'].extend(quality_assessment['quality_issues'])
            if quality_assessment['quality_issues']:
                logger.info(f"Quality issues: {', '.join(quality_assessment['quality_issues'])}")

        except Exception as e:
            logger.error(f"Image conversion failed: {e}")
    else:
        # For image files (JPG, PNG, etc.)
        try:
            import cv2
            import numpy as np
            img = cv2.imread(file_path)
            if img is not None:
                if rotation_angle != 0:
                    try:
                        from PIL import Image
                        import io
                        pil_img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
                        pil_img_rotated = pil_img.rotate(rotation_angle, expand=True, fillcolor='white')
                        img = cv2.cvtColor(np.array(pil_img_rotated), cv2.COLOR_RGB2BGR)
                        logger.info(f"✓ Applied cached rotation ({rotation_angle}°) to full page read")
                    except Exception as e:
                        logger.warning(f"⚠ Warning applying rotation to full page: {e}")

                _, buf = cv2.imencode('.jpg', img)
                img_bytes = buf.tobytes()
                img_bytes, was_downsampled, orig_size, new_size = _downsample_high_res_image(img_bytes, MAX_IMAGE_DIMENSION)
                quality_metadata['image_resolution'] = f"{orig_size[0]}x{orig_size[1]}"

                if was_downsampled:
                    logger.info(f"Downsampled: {orig_size[0]}x{orig_size[1]}  {new_size[0]}x{new_size[1]}")
                    quality_metadata['quality_issues'].append("רזולוציה גבוהה - הופחתה אוטומטית")
                elif max(orig_size) > WARN_IMAGE_DIMENSION:
                    logger.info(f"High resolution: {orig_size[0]}x{orig_size[1]}")
                    quality_metadata['quality_issues'].append("רזולוציה גבוהה")

                quality_assessment = _assess_image_quality(img_bytes)
                if quality_assessment['quality_issues']:
                    quality_metadata['quality_issues'].extend(quality_assessment['quality_issues'])
                if quality_assessment['quality_issues']:
                    logger.info(f"Quality issues: {', '.join(quality_assessment['quality_issues'])}")
            else:
                with open(file_path, "rb") as f:
                    img_bytes = f.read()
                img_bytes, was_downsampled, orig_size, new_size = _downsample_high_res_image(img_bytes, MAX_IMAGE_DIMENSION)
                quality_metadata['image_resolution'] = f"{orig_size[0]}x{orig_size[1]}"
                if was_downsampled:
                    logger.info(f"Downsampled: {orig_size[0]}x{orig_size[1]}  {new_size[0]}x{new_size[1]}")
                    quality_metadata['quality_issues'].append("רזולוציה גבוהה - הופחתה אוטומטית")
                quality_assessment = _assess_image_quality(img_bytes)
                if quality_assessment['quality_issues']:
                    quality_metadata['quality_issues'].extend(quality_assessment['quality_issues'])
                if quality_assessment['quality_issues']:
                    logger.info(f"Quality issues: {', '.join(quality_assessment['quality_issues'])}")
        except Exception as e:
            logger.error(f"cv2 failed, trying direct read: {e}")
            try:
                with open(file_path, "rb") as f:
                    img_bytes = f.read()
                img_bytes, was_downsampled, orig_size, new_size = _downsample_high_res_image(img_bytes, MAX_IMAGE_DIMENSION)
                quality_metadata['image_resolution'] = f"{orig_size[0]}x{orig_size[1]}"
                if was_downsampled:
                    logger.info(f"Downsampled: {orig_size[0]}x{orig_size[1]}  {new_size[0]}x{new_size[1]}")
                    quality_metadata['quality_issues'].append("רזולוציה גבוהה - הופחתה אוטומטית")
                quality_assessment = _assess_image_quality(img_bytes)
                if quality_assessment['quality_issues']:
                    quality_metadata['quality_issues'].extend(quality_assessment['quality_issues'])
                if quality_assessment['quality_issues']:
                    logger.info(f"Quality issues: {', '.join(quality_assessment['quality_issues'])}")
            except Exception as e2:
                logger.error(f"Error reading image: {e2}")
                return None, None, None

    if not img_bytes:
        logger.info("No image data")
        return None, None, None

    # ── Enhance contrast ─────────────────────────────────────────
    if used_high_res and title_block_bytes:
        title_block_bytes, was_enhanced, enhancement_metrics = _enhance_contrast_for_title_block(title_block_bytes)
        if was_enhanced:
            quality_metadata['quality_issues'].append(
                f"Enhanced TB: brightness {enhancement_metrics['original_brightness']:.0f}"
                f"→{enhancement_metrics['enhanced_brightness']:.0f}"
            )
    else:
        img_bytes, was_enhanced, enhancement_metrics = _enhance_contrast_for_title_block(img_bytes)
        if was_enhanced:
            quality_metadata['quality_issues'].append(
                f"Enhanced: brightness {enhancement_metrics['original_brightness']:.0f}"
                f"→{enhancement_metrics['enhanced_brightness']:.0f}"
            )

    # ── Step 2: OCR extraction ───────────────────────────────────
    combined_text_stage1 = ""
    combined_text_stage2 = ""
    combined_text_stage3 = ""

    if used_high_res and title_block_bytes:
        logger.info(f"Running OCR on high-res TITLE BLOCK...")
        ocr_results_stage1 = _run_with_timeout(
            ocr_engine.extract_all,
            args=(title_block_bytes,),
            timeout=300,
            stage_name=f"OCR Title Block ({Path(file_path).name})"
        )
        if ocr_results_stage1 is None:
            ocr_results_stage1 = {}
        image_b64_stage1 = base64.b64encode(title_block_bytes).decode('utf-8')
        combined_text_stage1 = ocr_engine.combine_results(ocr_results_stage1)

        _check_timeout()

        logger.info(f"Running OCR on OVERVIEW for other stages...")
        if not timeout_triggered:
            ocr_results_stage2 = _run_with_timeout(
                ocr_engine.extract_all,
                args=(img_bytes,),
                timeout=300,
                stage_name=f"OCR Overview ({Path(file_path).name})"
            )
            if ocr_results_stage2 is None:
                ocr_results_stage2 = {}
        else:
            ocr_results_stage2 = {}
        image_b64_stage2 = base64.b64encode(img_bytes).decode('utf-8')
        combined_text_stage2 = ocr_engine.combine_results(ocr_results_stage2)

        image_b64_stage3 = image_b64_stage2
        combined_text_stage3 = combined_text_stage2
    else:
        logger.info(f"Running OCR on FULL image...")
        ocr_results_stage1 = _run_with_timeout(
            ocr_engine.extract_all,
            args=(img_bytes,),
            timeout=300,
            stage_name=f"OCR Full ({Path(file_path).name})"
        )
        if ocr_results_stage1 is None:
            ocr_results_stage1 = {}
        image_b64_stage1 = base64.b64encode(img_bytes).decode('utf-8')
        combined_text_stage1 = ocr_engine.combine_results(ocr_results_stage1)

        image_b64_stage2 = image_b64_stage1
        combined_text_stage2 = combined_text_stage1
        image_b64_stage3 = image_b64_stage1
        combined_text_stage3 = combined_text_stage1

    # Combine with pdfplumber text
    if text_snippet:
        full_text_stage1 = f"{text_snippet[:1500]}\n\n{combined_text_stage1}"
        full_text_stage2 = f"{text_snippet}\n\n{combined_text_stage2}"
        full_text_stage3 = f"{text_snippet}\n\n{combined_text_stage3}"
    else:
        full_text_stage1 = combined_text_stage1
        full_text_stage2 = combined_text_stage2
        full_text_stage3 = combined_text_stage3

    logger.info(f"Sending to GPT ({sum(selected_stages.values())} stages selected)")

    # ── Customer detection ───────────────────────────────────────
    is_rafael_customer = False
    is_iai_customer = False

    # Combine OCR + pdfplumber for more reliable detection
    text_upper = combined_text_stage1.upper()
    if _pdfplumber_text:
        text_upper = text_upper + "\n" + _pdfplumber_text.upper()

    # Check for IAI patterns FIRST
    iai_indicators = 0
    if any(term in text_upper for term in ['IAI', 'I.A.I', 'ELTA', 'תעשייה אווירית', 'תעשיה אווירית', 'AEROSPACE INDUSTRIES']):
        iai_indicators += 3
    if 'INDUSTRIES' in text_upper and 'AEROSPACE' in text_upper:
        iai_indicators += 2
    if 'SYSTEM MISSILE' in text_upper or 'SPACE GROUP' in text_upper:
        iai_indicators += 2
    if '1933A' in text_upper or '1934A' in text_upper:
        iai_indicators += 2
    if 'PROJ ID' in text_upper or 'PROJECT ID' in text_upper:
        iai_indicators += 1

    if iai_indicators > 0:
        is_iai_customer = True
        logger.info(f"🎯 IAI Model (OCR, score: {iai_indicators})")
    else:
        rafael_indicators = 0
        if 'RAFAEL' in text_upper:
            rafael_indicators += 2
        if 'COPYRIGHT' in text_upper and 'RAFAEL' in text_upper:
            rafael_indicators += 2
        if any(part_prefix in text_upper for part_prefix in ['RF21', 'RF20', 'RF19', 'RF18', 'RF17']):
            rafael_indicators += 1
        # PRODUCTION ROUTING CHART is specific to RAFAEL variant format
        if 'PRODUCTION ROUTING CHART' in text_upper:
            rafael_indicators += 2

        if rafael_indicators >= 3:
            is_rafael_customer = True
            logger.info(f"🎯 RAFAEL Model (OCR, score: {rafael_indicators})")
        else:
            # Vision-based customer detection
            try:
                messages = [{"role": "user", "content": [
                    {"type": "text", "text": "Look at the title block/logo. Who is the manufacturer? Return JSON: {\"manufacturer\": \"RAFAEL\" | \"IAI\" | \"OTHER\"}"},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64_stage1}"}}
                ]}]
                response = _call_vision_api_with_retry(client, messages, max_tokens=50, temperature=0, stage_num=STAGE_CLASSIFICATION)
                if response:
                    content = response.choices[0].message.content
                    try:
                        data = json.loads(content)
                        answer = data.get("manufacturer", "").upper()
                    except Exception as e:
                        logger.debug(f"[CUSTOMER_DETECTION_JSON_PARSE] Error: {e}")
                        answer = content.upper()
                    if 'IAI' in answer and 'RAFAEL' not in answer:
                        is_iai_customer = True
                        logger.info(f"🎯 IAI Model (Vision)")
                    elif 'RAFAEL' in answer and 'IAI' not in answer:
                        is_rafael_customer = True
                        logger.info(f"🎯 RAFAEL Model (Vision)")
            except Exception as e:
                logger.error(f"Customer detection failed: {e}")

    if not is_rafael_customer and not is_iai_customer:
        logger.info(f"📋 Standard Customer")

    # ── Choose function set ──────────────────────────────────────
    if is_rafael_customer:
        extract_basic_fn = extract_basic_info_rafael
        extract_processes_fn = extract_processes_info_rafael
        extract_notes_fn = extract_notes_text_rafael
        extract_area_fn = extract_area_info_rafael
        identify_layout_fn = identify_drawing_layout_rafael
    elif is_iai_customer:
        extract_basic_fn = extract_basic_info_iai
        extract_processes_fn = extract_processes_info_iai
        extract_notes_fn = extract_notes_text_iai
        extract_area_fn = extract_area_info_iai
        identify_layout_fn = identify_drawing_layout_iai
    else:
        extract_basic_fn = extract_basic_info
        extract_processes_fn = extract_processes_info
        extract_notes_fn = extract_notes_text
        extract_area_fn = calculate_geometric_area
        identify_layout_fn = identify_drawing_layout

    # ── Stage 0.5: Layout identification ─────────────────────────
    tokens_in_layout, tokens_out_layout = 0, 0
    drawing_layout = None

    drawing_layout, tokens_in_layout, tokens_out_layout = identify_layout_fn(image_b64_stage1, client)
    if drawing_layout:
        logger.info(f"Layout: {drawing_layout[:80]}...")

    # ── Stage 1: Basic Info (Title Block) ────────────────────────
    tokens_in_1, tokens_out_1 = 0, 0
    basic_data: Dict = {}
    retry_log = ""

    stage1_enable_retry = enable_retry
    try:
        import cv2
        import numpy as np
        nparr = np.frombuffer(img_bytes, np.uint8)
        img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        if img_cv is not None:
            image_width = img_cv.shape[1]
            if stage1_skip_retry_resolution_px > 0 and image_width >= stage1_skip_retry_resolution_px:
                logger.info(f"ℹ️ High resolution ({image_width}px width, threshold {stage1_skip_retry_resolution_px}px) - skipping attempts 2 & 3 to save time")
                stage1_enable_retry = False
    except Exception as e:
        logger.debug(f"[STAGE1_RETRY_RESOLUTION_CHECK] Error: {e}")

    if selected_stages.get(1, False):
        filename = Path(file_path).stem
        stage1_result = _run_with_timeout(
            extract_stage1_with_retry,
            kwargs=dict(
                img_bytes=img_bytes,
                filename=filename,
                extract_basic_fn=extract_basic_fn,
                client=client,
                ocr_engine=ocr_engine,
                is_rafael=is_rafael_customer,
                is_iai=is_iai_customer,
                used_high_res=used_high_res,
                title_block_bytes=title_block_bytes,
                enable_retry=stage1_enable_retry,
                pdfplumber_text=_pdfplumber_text,
            ),
            timeout=STAGE_HARD_TIMEOUT,
            stage_name=f"Stage 1 ({Path(file_path).name})"
        )
        if stage1_result and isinstance(stage1_result, tuple) and len(stage1_result) >= 4:
            basic_data, tokens_in_1, tokens_out_1, retry_log = stage1_result
        else:
            basic_data, tokens_in_1, tokens_out_1, retry_log = {}, 0, 0, ""

        if not basic_data:
            basic_data = {}
            logger.error(f"Stage 1: Failed after retries - continuing with other stages")
        else:
            logger.info(f"Stage 1: {list(basic_data.keys())}")
    else:
        logger.info(f"Stage 1: Skipped")

    _check_timeout()

    # Check if GUI requested skip after Stage 1
    if _ocr_mod._gui_should_skip and _ocr_mod._gui_should_skip():
        logger.info("Skip requested - stopping file processing")
        return None, None, None

    # ── Stage 2: Processes Info ──────────────────────────────────
    tokens_in_2, tokens_out_2 = 0, 0
    process_data: Dict = {}
    notes_text_already_extracted = None

    if selected_stages.get(2, False) and not timeout_triggered:
        max_stage2_retries = 2
        attempt = 0
        while attempt <= max_stage2_retries:
            logger.debug(f"      DEBUG Stage 2: Calling extract_processes_fn (attempt {attempt + 1}/{max_stage2_retries + 1})...")
            result = _run_with_timeout(
                extract_processes_fn,
                args=(full_text_stage2, image_b64_stage2, client),
                kwargs=dict(additional_images=additional_pages_b64),
                timeout=300,
                stage_name=f"Stage 2 GPT Vision ({Path(file_path).name})"
            )
            logger.debug(f"      DEBUG Stage 2: Result returned - type: {type(result)}, is tuple: {isinstance(result, tuple)}")

            if result and isinstance(result, tuple) and len(result) >= 3:
                extracted_data, tokens_in_2, tokens_out_2 = result[0], result[1], result[2]
                logger.debug(f"      DEBUG Stage 2: Extracted data type: {type(extracted_data)}, is dict: {isinstance(extracted_data, dict)}")
            else:
                extracted_data = None
                tokens_in_2, tokens_out_2 = 0, 0
                logger.debug(f"      DEBUG Stage 2: Result validation failed - tuple check: {isinstance(result, tuple)}, len: {len(result) if isinstance(result, tuple) else 'N/A'}")

            if extracted_data and isinstance(extracted_data, dict):
                process_data = extracted_data
                logger.info(f"Stage 2: SUCCESS - {list(process_data.keys())}")
                break
            else:
                logger.debug(f"      DEBUG Stage 2: Attempt {attempt + 1} failed - extracted_data is dict: {isinstance(extracted_data, dict)}")

            attempt += 1
            if attempt <= max_stage2_retries:
                logger.error(f"Stage 2: Failed - retrying ({attempt}/{max_stage2_retries})")

        if not process_data:
            process_data = {}
            if selected_stages.get(5, False):
                logger.error(f"Stage 2: Failed - attempting Stage 5 fallback from NOTES...")
                if not selected_stages.get(3, False):
                    logger.info(f"Stage 5: Requires Stage 3 enabled for NOTES extraction - skipping")
                else:
                    logger.info(f"Running Stage 3 (NOTES) first for Stage 5 fallback...")
                    result = _run_with_timeout(
                        extract_notes_fn,
                        args=(full_text_stage3, image_b64_stage3, client),
                        kwargs=dict(additional_images=additional_pages_b64),
                        timeout=300,
                        stage_name=f"Stage 3/5 GPT NOTES ({Path(file_path).name})"
                    )
                    notes_text_already_extracted = None
                    if result and isinstance(result, tuple) and len(result) > 0:
                        notes_text_already_extracted = result[0]

                    if notes_text_already_extracted and isinstance(notes_text_already_extracted, str):
                        logger.info(f"Stage 3 (for S5): NOTES extracted ({len(notes_text_already_extracted)} chars)")

                        validated_notes, validation_report = validate_notes_before_stage5(
                            notes_text_already_extracted,
                            debug_prefix="         "
                        )

                        if validated_notes is None:
                            logger.error(f"Stage 5: NOTES validation failed - skipping")
                        else:
                            result_s5 = extract_processes_from_notes(validated_notes, client)

                            if result_s5 and isinstance(result_s5, tuple) and len(result_s5) >= 3:
                                s5_data, s5_tokens_in, s5_tokens_out = result_s5[0], result_s5[1], result_s5[2]
                            else:
                                s5_data, s5_tokens_in, s5_tokens_out = None, 0, 0

                            if s5_data and isinstance(s5_data, dict):
                                process_data = s5_data
                                tokens_in_2 = s5_tokens_in
                                tokens_out_2 = s5_tokens_out
                                logger.info(f"Stage 5 (Fallback): SUCCESS - {list(process_data.keys())}")
                            else:
                                logger.error(f"Stage 5: Failed - could not extract processes from notes")
                    else:
                        logger.info(f"Stage 5: Could not extract NOTES from Stage 3 - skipping (notes_text_already_extracted: {type(notes_text_already_extracted).__name__})")
            else:
                logger.error(f"Stage 2: Failed - continuing with other stages")
    else:
        logger.info(f"Stage 2: Skipped")

    _check_timeout()

    if _ocr_mod._gui_should_skip and _ocr_mod._gui_should_skip():
        logger.info("Skip requested - stopping file processing")
        return None, None, None

    # ── Stage 3: NOTES Text Only ─────────────────────────────────
    tokens_in_3, tokens_out_3 = 0, 0
    notes_text = None

    if selected_stages.get(3, False) and not timeout_triggered:
        if notes_text_already_extracted:
            notes_text = notes_text_already_extracted
            logger.info(f"Stage 3: NOTES extraction (reused from Stage 5 fallback)")
        else:
            stage3_result = _run_with_timeout(
                extract_notes_fn,
                args=(full_text_stage3, image_b64_stage3, client),
                kwargs=dict(additional_images=additional_pages_b64),
                timeout=300,
                stage_name=f"Stage 3 GPT NOTES ({Path(file_path).name})"
            )
            if stage3_result and isinstance(stage3_result, tuple) and len(stage3_result) >= 3:
                notes_text, tokens_in_3, tokens_out_3 = stage3_result[0], stage3_result[1], stage3_result[2]
            else:
                notes_text, tokens_in_3, tokens_out_3 = None, 0, 0
            if notes_text:
                logger.info(f"Stage 3: NOTES extraction")
                logger.info(f"🔍 Validating extracted NOTES...")
                validated_notes, validation_report = validate_notes_before_stage5(
                    notes_text,
                    debug_prefix="         "
                )
                if validated_notes is None:
                    logger.error(f"⚠️ NOTES failed validation - will not be used for Stage 5")
                    notes_text = None
            else:
                logger.info(f"Stage 3: No notes found - continuing")
    else:
        logger.info(f"Stage 3: Skipped")

    _check_timeout()

    if _ocr_mod._gui_should_skip and _ocr_mod._gui_should_skip():
        logger.info("Skip requested - stopping file processing")
        return None, None, None

    # ── Stage 4: Geometric Area ──────────────────────────────────
    tokens_in_4, tokens_out_4 = 0, 0
    geometric_area = None

    if selected_stages.get(4, False) and not timeout_triggered:
        stage4_result = _run_with_timeout(
            extract_area_fn,
            args=(full_text_stage2, image_b64_stage2, client),
            timeout=300,
            stage_name=f"Stage 4 Area ({Path(file_path).name})"
        )
        if stage4_result and isinstance(stage4_result, tuple) and len(stage4_result) >= 3:
            geometric_area, tokens_in_4, tokens_out_4 = stage4_result[0], stage4_result[1], stage4_result[2]
        else:
            geometric_area, tokens_in_4, tokens_out_4 = None, 0, 0

        if geometric_area:
            logger.info(f"Stage 4: Geometric area calculation")
        else:
            logger.info(f"Stage 4: Could not estimate area - continuing")
    else:
        logger.info(f"Stage 4: Skipped")

    if _ocr_mod._gui_should_skip and _ocr_mod._gui_should_skip():
        logger.info("Skip requested - stopping file processing")
        return None, None, None

    # ── Merge results ────────────────────────────────────────────
    if not isinstance(basic_data, dict):
        basic_data = {}
    if not isinstance(process_data, dict):
        process_data = {}

    result_data = {**basic_data, **process_data, "notes_full_text": notes_text}

    if retry_log:
        result_data['stage1_retry_log'] = retry_log

    if timeout_triggered:
        existing = str(result_data.get('validation_warnings', '') or '').strip()
        timeout_msg = f"TIMEOUT after {timeout_elapsed:.0f}s"
        result_data['validation_warnings'] = f"{existing} | {timeout_msg}" if existing else timeout_msg

    # Quality metadata
    result_data['image_resolution'] = quality_metadata['image_resolution']
    result_data['quality_issues'] = ', '.join(quality_metadata['quality_issues']) if quality_metadata['quality_issues'] else ''

    if drawing_layout:
        result_data['drawing_layout'] = drawing_layout

    if geometric_area:
        result_data['part_area'] = geometric_area

    # ── Post-process: MASKING + INSERTS from NOTES ───────────────
    result_data = post_process_summary_from_notes(
        result_data,
        pdfplumber_text=_pdfplumber_text or '',
        ocr_text=combined_text_stage2,
    )

    # ── Multi-strategy P.N. extraction & voting ──────────────────
    filename = Path(file_path).stem

    vision_pn = str(result_data.get('part_number', '')).strip()
    vision_dn = str(result_data.get('drawing_number', '')).strip()

    pdf_extracted = extract_pn_dn_from_text(_pdfplumber_text)
    pdfplumber_pn = pdf_extracted.get('part_number', '')
    pdfplumber_dn = pdf_extracted.get('drawing_number', '')

    tesseract_pn = ''
    tesseract_dn = ''
    if combined_text_stage1:
        tess_extracted = extract_pn_dn_from_text(combined_text_stage1)
        tesseract_pn = tess_extracted.get('part_number', '')
        tesseract_dn = tess_extracted.get('drawing_number', '')

    # ── Guard: When RAFAEL explicitly found separate P.N. field, don't let
    #    the vote replace it with a value that equals the Drawing Number.
    #    pdfplumber regex often confuses PN/DN fields in RAFAEL title blocks.
    _searched_pn = result_data.get('_searched_for_pn_field', False)

    if vision_pn != 'N/A' and pdfplumber_pn:
        best_pn, pn_source = vote_best_pn(vision_pn, pdfplumber_pn, tesseract_pn, filename)
        if best_pn != vision_pn:
            # If RAFAEL extraction explicitly found separate P.N./DN fields,
            # reject vote winner if it equals the drawing number (= field confusion)
            dn_norm = re.sub(r'[^A-Za-z0-9]', '', vision_dn).upper()
            best_norm = re.sub(r'[^A-Za-z0-9]', '', best_pn).upper()
            if _searched_pn and dn_norm and best_norm == dn_norm:
                logger.info(
                    f"\U0001f5f3\ufe0f P.N. vote: Winner '{best_pn}' == Drawing# '{vision_dn}' "
                    f"— rejected (RAFAEL explicit P.N. field). Keeping Vision: '{vision_pn}'"
                )
            else:
                logger.info(f"\U0001f5f3\ufe0f P.N. vote: Vision='{vision_pn}' pdfplumber='{pdfplumber_pn}' → Winner: '{best_pn}' ({pn_source})")
                result_data['part_number'] = best_pn
    elif vision_pn == 'N/A' and pdfplumber_pn:
        logger.info(f"\U0001f5f3\ufe0f P.N. Vision=N/A, pdfplumber='{pdfplumber_pn}' → using pdfplumber")
        result_data['part_number'] = pdfplumber_pn

    if vision_dn and pdfplumber_dn:
        best_dn, dn_source = vote_best_pn(vision_dn, pdfplumber_dn, tesseract_dn, filename)
        if best_dn != vision_dn:
            logger.info(f"\U0001f5f3\ufe0f Drawing# vote: Vision='{vision_dn}' pdfplumber='{pdfplumber_dn}' → Winner: '{best_dn}' ({dn_source})")
            result_data['drawing_number'] = best_dn

    # ── Token totals ─────────────────────────────────────────────
    total_input = tokens_in_1 + tokens_in_2 + tokens_in_3 + tokens_in_4 + tokens_in_layout
    total_output = tokens_out_1 + tokens_out_2 + tokens_out_3 + tokens_out_4 + tokens_out_layout

    # ── Per-stage accurate cost (uses real per-model prices) ─────
    _stage_tok = {
        0: (tokens_in_layout, tokens_out_layout),
        1: (tokens_in_1, tokens_out_1),
        2: (tokens_in_2, tokens_out_2),
        3: (tokens_in_3, tokens_out_3),
        4: (tokens_in_4, tokens_out_4),
    }
    pipeline_cost_usd = sum(
        _calculate_stage_cost(ti, to, s) for s, (ti, to) in _stage_tok.items()
    )

    # ── Context for caller (sanity checks / confidence) ──────────
    context = {
        'pdfplumber_text': _pdfplumber_text,
        'combined_text_stage1': combined_text_stage1,
        'is_rafael': is_rafael_customer,
        'is_iai': is_iai_customer,
        'num_pages': num_pages,
        'pipeline_cost_usd': pipeline_cost_usd,
    }

    return result_data, (total_input, total_output), context
