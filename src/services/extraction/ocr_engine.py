"""
OCR Engine and Stage 1 Retry Logic
===================================
Extracted from customer_extractor_v3_dual.py
"""
import os
import re
import base64
import cv2
import numpy as np
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Any
from PIL import Image

from openai import AzureOpenAI
from src.utils.logger import get_logger
from src.services.ai.vision_api import (
    _resolve_stage_call_config,
    _calculate_stage_cost,
    _chat_create_with_token_compat,
    _call_vision_api_with_retry,
)
from src.services.image.processing import (
    _downsample_high_res_image,
    _enhance_contrast_for_title_block,
    _extract_image_smart,
    _assess_image_quality,
)
from src.services.extraction.filename_utils import (
    check_value_in_filename,
    _disambiguate_part_number,
)

logger = get_logger(__name__)

DEBUG_ENABLED = os.getenv("AI_DRAW_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}
IAI_TOP_RED_FALLBACK_ENABLED = os.getenv("IAI_TOP_RED_FALLBACK_ENABLED", "true").strip().lower() in {"1", "true", "yes", "on"}

_gui_should_stop = None
_gui_should_skip = None


def set_gui_callbacks(stop_fn, skip_fn) -> None:
    global _gui_should_stop, _gui_should_skip
    _gui_should_stop = stop_fn
    _gui_should_skip = skip_fn


def debug_print(message: str) -> None:
    if DEBUG_ENABLED:
        print(message)


class MultiOCREngine:
    """מנוע OCR עם Tesseract"""
    
    def __init__(self) -> None:
        self._cache: Dict[str, Dict[str, str]] = {}  # md5 -> OCR result
        self._cache_hits = 0
        self._cache_misses = 0
    
    def extract_title_block_region(self, img_bytes: bytes, enlarge_factor: float = 2.0) -> Tuple[Optional[bytes], str]:
        """
        חיתוך אזור Title Block והגדלתו לקריאה טובה יותר
        
        Returns:
            (cropped_enlarged_bytes, location_description)
        """
        try:
            nparr = np.frombuffer(img_bytes, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            
            if img is None:
                return img_bytes, "full"
            
            height, width = img.shape[:2]
            
            # Try multiple title block locations (RAFAEL uses different positions)
            # We'll try bottom strip first (covers both left and right)
            locations = [
                # Full Bottom (safest - covers both left and right title blocks)
                {
                    'name': 'bottom-full',
                    'y_start': int(height * 0.75),
                    'y_end': height,
                    'x_start': 0,
                    'x_end': width
                }
            ]
            
            # Use full bottom strip to catch title block anywhere
            location = locations[0]
            
            title_block = img[location['y_start']:location['y_end'], 
                            location['x_start']:location['x_end']]
            
            # Enlarge the title block for better OCR/Vision
            if enlarge_factor > 1.0:
                new_width = int(title_block.shape[1] * enlarge_factor)
                new_height = int(title_block.shape[0] * enlarge_factor)
                
                # Check if enlarged size exceeds OpenCV's limit (65500 pixels)
                max_dimension = max(new_width, new_height)
                if max_dimension > 60000:  # Leave safety margin
                    # Reduce enlargement factor to fit within limit
                    scale_down = 60000 / max_dimension
                    enlarge_factor = enlarge_factor * scale_down
                    new_width = int(title_block.shape[1] * enlarge_factor)
                    new_height = int(title_block.shape[0] * enlarge_factor)
                    logger.info(f"Reduced enlargement to {enlarge_factor:.1f}x to avoid OpenCV size limit")
                
                title_block = cv2.resize(title_block, (new_width, new_height), 
                                        interpolation=cv2.INTER_CUBIC)
            
            # Encode as JPEG with high quality to preserve text clarity
            # Quality 95 is essential for accurate character recognition
            success, encoded = cv2.imencode('.jpg', title_block, [cv2.IMWRITE_JPEG_QUALITY, 95])
            
            if success:
                img_bytes_result = encoded.tobytes()
                size_mb = len(img_bytes_result) / (1024 * 1024)
                
                # If still too large, reduce quality slightly (but not too much)
                if size_mb > 18:  # Keep under 18MB to be safe (limit is 20MB)
                    logger.info(f"Image too large ({size_mb:.1f}MB), reducing quality to 85...")
                    success, encoded = cv2.imencode('.jpg', title_block, [cv2.IMWRITE_JPEG_QUALITY, 85])
                    if success:
                        img_bytes_result = encoded.tobytes()
                        size_mb = len(img_bytes_result) / (1024 * 1024)
                
                logger.info(f"Extracted Title Block ({location['name']}): "
                      f"{title_block.shape[1]}x{title_block.shape[0]}px "
                      f"(enlarged {enlarge_factor:.1f}x, {size_mb:.1f}MB)")
                return img_bytes_result, location['name']
            
            return img_bytes, "full"
            
        except Exception as e:
            logger.error(f"Title Block extraction failed: {e}")
            return img_bytes, "full"
    
    def preprocess_image(self, img_bytes: bytes) -> np.ndarray:
        """עיבוד מקדים לשיפור OCR"""
        nparr = np.frombuffer(img_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        
        if img is None:
            return None
        
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        denoised = cv2.fastNlMeansDenoising(enhanced)
        thresh = cv2.adaptiveThreshold(
            denoised, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 11, 2
        )
        
        return thresh
    
    def detect_title_block(self, img_bytes: bytes) -> Tuple[Optional[bytes], bool]:
        """זיהוי אזור Title Block - Bottom-Right"""
        try:
            nparr = np.frombuffer(img_bytes, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            
            if img is None:
                return img_bytes, True
            
            height, width = img.shape[:2]
            
            # Bottom-right corner (30% height, 40% width)
            y_start = int(height * 0.7)
            x_start = int(width * 0.6)
            
            title_block_region = img[y_start:height, x_start:width]
            
            # Check if region is too large
            region_pixels = title_block_region.shape[0] * title_block_region.shape[1]
            original_pixels = height * width
            
            if region_pixels > original_pixels * 0.7:
                logger.info(f"Using full image (title block region too large)")
                return img_bytes, True
            
            success, encoded = cv2.imencode('.png', title_block_region)
            if success:
                logger.info(f"Title Block region: {title_block_region.shape[1]}x{title_block_region.shape[0]}px (bottom-right)")
                return encoded.tobytes(), False
            
            return img_bytes, True
            
        except Exception as e:
            logger.error(f"Title Block detection failed: {e}")
            return img_bytes, True
    
    @staticmethod
    def _safe_downscale_for_tesseract(img_array: np.ndarray, max_pixels: int = 30_000_000) -> np.ndarray:
        """Downscale image if it exceeds Tesseract's pixel limit (~32M)."""
        h, w = img_array.shape[:2]
        total = h * w
        if total > max_pixels:
            scale = (max_pixels / total) ** 0.5
            new_w, new_h = int(w * scale), int(h * scale)
            logger.info(f"Downscaling for Tesseract: {w}x{h} → {new_w}x{new_h} ({total/1e6:.0f}M → {new_w*new_h/1e6:.0f}M px)")
            img_array = cv2.resize(img_array, (new_w, new_h), interpolation=cv2.INTER_AREA)
        return img_array

    def extract_with_tesseract(self, img_bytes: bytes) -> Optional[str]:
        """OCR עם Tesseract"""
        try:
            import pytesseract
            from PIL import Image
            import io
            
            tess_path = os.getenv("TESSERACT_PATH")
            if tess_path:
                pytesseract.pytesseract.tesseract_cmd = tess_path
            
            processed = self.preprocess_image(img_bytes)
            
            if processed is not None:
                processed = self._safe_downscale_for_tesseract(processed)
                config = '--oem 3 --psm 6'
                lang = os.getenv("TESSERACT_LANGS", "eng+heb")
                text = pytesseract.image_to_string(processed, lang=lang, config=config)
            else:
                img = Image.open(io.BytesIO(img_bytes))
                # Check PIL image size too
                w, h = img.size
                if w * h > 30_000_000:
                    scale = (30_000_000 / (w * h)) ** 0.5
                    img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
                    logger.info(f"Downscaled PIL image for Tesseract: {w}x{h} → {img.size[0]}x{img.size[1]}")
                lang = os.getenv("TESSERACT_LANGS", "eng+heb")
                text = pytesseract.image_to_string(img, lang=lang)
            
            return text
        
        except Exception as e:
            logger.error(f"Tesseract failed: {e}")
            return None
    
    def extract_all(self, img_bytes: bytes) -> Dict[str, str]:
        """מריץ OCR עם Tesseract (with caching)"""
        import hashlib
        cache_key = hashlib.md5(img_bytes).hexdigest()

        if cache_key in self._cache:
            self._cache_hits += 1
            logger.info(f"OCR cache hit (saved ~5-10 sec)")
            return self._cache[cache_key]

        self._cache_misses += 1

        results = {}
        
        tess = self.extract_with_tesseract(img_bytes)
        if tess:
            results['tesseract'] = tess
            logger.info(f"Tesseract: {len(tess)} chars")

        self._cache[cache_key] = results
        return results

    def get_cache_stats(self) -> str:
        """Return cache statistics."""
        total = self._cache_hits + self._cache_misses
        if total == 0:
            return "OCR cache: no calls"
        hit_rate = (self._cache_hits / total) * 100
        return (f"OCR cache: {self._cache_hits} hits, {self._cache_misses} misses "
                f"({hit_rate:.0f}% hit rate, ~{self._cache_hits * 8}s saved)")
    
    def combine_results(self, results: Dict[str, str]) -> str:
        """משלב תוצאות OCR"""
        combined = []
        
        for engine, text in results.items():
            if text and text.strip():
                combined.append(f"=== {engine.upper()} ===")
                combined.append(text[:3000])
                combined.append("")
        
        return "\n".join(combined)


# 
# פונקציות עזר
# 

def extract_stage1_with_retry(
    img_bytes: bytes,
    filename: str,
    extract_basic_fn,
    client: AzureOpenAI,
    ocr_engine: MultiOCREngine,
    is_rafael: bool,
    is_iai: bool,
    used_high_res: bool,
    title_block_bytes: Optional[bytes],
    enable_retry: bool = False,
    pdfplumber_text: str = '',
    email_parts_hint: str = ''
) -> Tuple[Optional[Dict], int, int, str]:
    """
     Stage 1 with smart retry mechanism
    
    Smart attempts with progressive enlargement (when retry enabled):
    1. Original Title Block
    2. Enlarged region
    3. Larger region + pre-processing

    Stops early only if Part# found in filename (success)
    Can limit to attempt 1 only with enable_retry=False
    """
    MAX_RETRIES = 3 if enable_retry else 1
    ENLARGE_STEP = 0.20  # 20% per retry
    
    attempt_log = []
    
    for attempt in range(1, MAX_RETRIES + 1):
        logger.info(f"Attempt {attempt}/{MAX_RETRIES} for file: {filename}")
        
        # Determine which image to use
        if attempt == 1:
            # First attempt: use original
            if used_high_res and title_block_bytes:
                current_img_bytes = title_block_bytes
                ocr_results = ocr_engine.extract_all(current_img_bytes)
                image_b64 = base64.b64encode(current_img_bytes).decode('utf-8')
                full_text = ocr_engine.combine_results(ocr_results)
                source = "high-res TB"
            else:
                current_img_bytes = img_bytes
                ocr_results = ocr_engine.extract_all(current_img_bytes)
                image_b64 = base64.b64encode(current_img_bytes).decode('utf-8')
                full_text = ocr_engine.combine_results(ocr_results)
                source = "overview"
        
        elif attempt == 2:
            # Attempt 2: enlarge Title Block region by 20%
            if not used_high_res or not title_block_bytes:
                # Can't enlarge if not high-res - retry with same image
                source = f"retry {attempt}"
            else:
                # Fixed 20% enlargement for attempt 2 (from new 40%35% base)
                new_width_percent = 0.40 + ENLARGE_STEP   # 60%
                new_height_percent = 0.35 + ENLARGE_STEP  # 55%
                
                logger.info(f"Attempt {attempt}: Enlarging TB to {int(new_width_percent*100)}%{int(new_height_percent*100)}%...")
                
                # Re-extract Title Block with larger region
                try:
                    nparr = np.frombuffer(img_bytes, np.uint8)
                    img_full = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                    
                    if img_full is None:
                        source = f"retry {attempt}"
                    else:
                        full_height, full_width = img_full.shape[:2]
                        
                        # Calculate new region size
                        tb_height_new = int(full_height * new_height_percent)
                        tb_width_new = int(full_width * new_width_percent)
                        
                        # Extract larger region
                        tb_region_new = img_full[full_height - tb_height_new:, full_width - tb_width_new:]
                        
                        # Enlarge by 2x
                        tb_enlarged_new = cv2.resize(tb_region_new, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
                        
                        # Encode
                        _, tb_buffer_new = cv2.imencode('.png', tb_enlarged_new)
                        current_img_bytes = tb_buffer_new.tobytes()
                        
                        # OCR on new enlarged region
                        ocr_results = ocr_engine.extract_all(current_img_bytes)
                        image_b64 = base64.b64encode(current_img_bytes).decode('utf-8')
                        full_text = ocr_engine.combine_results(ocr_results)
                        
                        source = f"enlarged TB +20% (50%45%)"
                        
                except Exception as e:
                    logger.error(f"Failed to enlarge Title Block: {e}")
                    source = f"retry {attempt}"
        
        else:
            # Attempt 3: Pre-processing on enlarged image (40% enlargement = 70%65%)
            if not used_high_res or not title_block_bytes:
                # Can't enlarge if not high-res
                logger.info(f"Attempt {attempt}: Cannot enlarge (not in high-res mode)")
                break
            
            # 40% enlargement for attempt 3 (double the step from new base)
            new_width_percent = 0.40 + (ENLARGE_STEP * 2)   # 80%
            new_height_percent = 0.35 + (ENLARGE_STEP * 2)  # 75%
            
            logger.info(f"Attempt {attempt}: Enlarging TB to {int(new_width_percent*100)}%{int(new_height_percent*100)}% + Pre-processing...")
            
            # Re-extract Title Block with larger region
            try:
                nparr = np.frombuffer(img_bytes, np.uint8)
                img_full = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                
                if img_full is None:
                    break
                
                full_height, full_width = img_full.shape[:2]
                
                # Calculate new region size
                tb_height_new = int(full_height * new_height_percent)
                tb_width_new = int(full_width * new_width_percent)
                
                # Extract larger region
                tb_region_new = img_full[full_height - tb_height_new:, full_width - tb_width_new:]
                
                # Enlarge by 2x
                tb_enlarged_new = cv2.resize(tb_region_new, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
                
                #  Apply Pre-processing for attempt 3
                # Step 1: Denoise
                tb_enlarged_new = cv2.fastNlMeansDenoisingColored(
                    tb_enlarged_new, None, 
                    h=10, hColor=10, 
                    templateWindowSize=7, searchWindowSize=21
                )
                
                # Step 2: Sharpen
                kernel_sharp = np.array([[-1, -1, -1],
                                        [-1,  9, -1],
                                        [-1, -1, -1]])
                tb_enlarged_new = cv2.filter2D(tb_enlarged_new, -1, kernel_sharp)
                
                # Step 3: Adaptive Thresholding (for better text contrast)
                gray = cv2.cvtColor(tb_enlarged_new, cv2.COLOR_BGR2GRAY)
                thresh = cv2.adaptiveThreshold(
                    gray, 255, 
                    cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                    cv2.THRESH_BINARY, 11, 2
                )
                # Convert back to BGR for consistency
                tb_enlarged_new = cv2.cvtColor(thresh, cv2.COLOR_GRAY2BGR)
                
                # Encode
                _, tb_buffer_new = cv2.imencode('.png', tb_enlarged_new)
                current_img_bytes = tb_buffer_new.tobytes()
                
                # OCR on new enlarged region
                ocr_results = ocr_engine.extract_all(current_img_bytes)
                image_b64 = base64.b64encode(current_img_bytes).decode('utf-8')
                full_text = ocr_engine.combine_results(ocr_results)
                
                source = f"enlarged TB +40%+PP (70%65%)"
                
            except Exception as e:
                logger.error(f"Failed to enlarge Title Block: {e}")
                break
        
        # Run Stage 1 extraction
        if is_rafael:
            basic_data, tokens_in, tokens_out = extract_basic_fn(
                full_text, image_b64, client, ocr_engine,
                pdfplumber_text=pdfplumber_text,
                email_parts_hint=email_parts_hint,
                filename=filename,
            )
        elif is_iai:
            basic_data, tokens_in, tokens_out = extract_basic_fn(
                full_text,
                image_b64,
                client,
                ocr_engine,
                use_top_red_fallback=False
            )
        else:
            basic_data, tokens_in, tokens_out = extract_basic_fn(full_text, image_b64, client)
        
        if not basic_data:
            attempt_log.append(f"Attempt {attempt} ({source}): FAILED")
            continue
        
        # Check confidence by simulating filename check
        part_in_filename = check_value_in_filename(basic_data.get('part_number'), filename)
        drawing_in_filename = check_value_in_filename(basic_data.get('drawing_number'), filename)
        
        # Determine confidence
        if part_in_filename and drawing_in_filename:
            confidence = 'full'
        elif part_in_filename:
            confidence = 'high'
        elif drawing_in_filename:
            confidence = 'medium'
        else:
            confidence = 'low'
        
        # Log attempt result
        attempt_log.append(f"Attempt {attempt} ({source}): {confidence}")
        
        # If confidence is HIGH or FULL  success!
        if confidence in ['high', 'full']:
            log_str = "  ".join(attempt_log)
            if attempt > 1:
                logger.info(f"Success on attempt {attempt}")
            logger.info(f"SUCCESS on attempt {attempt} for '{filename}' with confidence '{confidence}'")
            return basic_data, tokens_in, tokens_out, log_str
        
        # Continue to next attempt if we haven't reached max retries
        if attempt < MAX_RETRIES:
            pass  # Continue silently to next attempt
    
    # All attempts exhausted - return last result
    log_str = "  ".join(attempt_log)
    # Ensure we always return a dict, never None
    result_dict = basic_data if basic_data and isinstance(basic_data, dict) else {}
    result_tokens_in = tokens_in if 'tokens_in' in locals() else 0
    result_tokens_out = tokens_out if 'tokens_out' in locals() else 0
    
    if not result_dict and attempt_log:
        logger.error(f"All {MAX_RETRIES} attempts failed: {' | '.join(attempt_log)}")

    # IAI-only fallback: use top red header after regular attempts when IDs are missing
    # OR when extracted IDs are low-confidence (do not match filename)
    if is_iai and IAI_TOP_RED_FALLBACK_ENABLED:
        part_num = str(result_dict.get('part_number', '') or '').strip()
        drawing_num = str(result_dict.get('drawing_number', '') or '').strip()
        part_in_filename = check_value_in_filename(part_num, filename)
        drawing_in_filename = check_value_in_filename(drawing_num, filename)

        missing_regular_ids = (not part_num) or (not drawing_num)
        low_confidence_regular_ids = (not part_in_filename) or (not drawing_in_filename)
        should_try_top_red = missing_regular_ids or low_confidence_regular_ids

        if should_try_top_red:
            try:
                logger.info("Stage 1B: IAI Top Red Header OCR Fallback")
                if missing_regular_ids:
                    logger.info("IAI fallback: regular attempts did not provide part/drawing. Trying top red header...")
                else:
                    logger.info("IAI fallback: regular attempts are low-confidence vs filename. Trying top red header...")

                full_ocr = ocr_engine.extract_all(img_bytes)
                full_text = ocr_engine.combine_results(full_ocr)
                full_image_b64 = base64.b64encode(img_bytes).decode('utf-8')

                fb_data, fb_in, fb_out = extract_basic_fn(
                    full_text,
                    full_image_b64,
                    client,
                    ocr_engine,
                    use_top_red_fallback=True
                )

                if fb_data and isinstance(fb_data, dict):
                    fb_part = str(fb_data.get('part_number', '') or '').strip()
                    fb_draw = str(fb_data.get('drawing_number', '') or '').strip()

                    if fb_part and fb_draw:
                        result_dict = fb_data
                        result_tokens_in += fb_in
                        result_tokens_out += fb_out
                        attempt_log.append("IAI top-red fallback: success")
                    else:
                        attempt_log.append("IAI top-red fallback: no valid identifier")
                else:
                    attempt_log.append("IAI top-red fallback: failed")
            except Exception as e:
                attempt_log.append(f"IAI top-red fallback error: {e}")
    elif is_iai and not IAI_TOP_RED_FALLBACK_ENABLED:
        attempt_log.append("IAI top-red fallback: disabled by env")
    
    log_str = "  ".join(attempt_log)
    return result_dict, result_tokens_in, result_tokens_out, log_str
