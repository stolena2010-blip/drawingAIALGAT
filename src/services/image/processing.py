"""
Image processing helpers for DrawingAI Pro.
=============================================

Contains image manipulation, quality assessment, rotation detection / correction,
and smart region-extraction logic extracted from customer_extractor_v3_dual.py.
"""

import base64
import io
import re
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import cv2
import numpy as np
from PIL import Image
from PIL.ExifTags import TAGS

# Raise PIL decompression bomb limit — enlarged title blocks can exceed
# the default 89M pixel threshold (e.g. 26194×4053 = 106M pixels).
Image.MAX_IMAGE_PIXELS = 300_000_000

from src.core.constants import (
    MAX_IMAGE_DIMENSION,
    STAGE_ROTATION,
    debug_print,
)
from src.utils.prompt_loader import load_prompt

# Late imports that live in the extractor today — we import them lazily inside
# the one function that needs them (_fix_image_rotation) to avoid circular deps.
# When those helpers are moved into their own module we will switch to a normal
# top-level import.

# ---------------------------------------------------------------------------
# Rotation detection cache (shared across calls within the same process)
# ---------------------------------------------------------------------------
_rotation_cache: Dict[str, tuple] = {}


# =========================================================================
# Image helpers
# =========================================================================

def _downsample_high_res_image(
    img_bytes: bytes,
    max_dimension: int = MAX_IMAGE_DIMENSION,
) -> Tuple[bytes, bool, Tuple[int, int], Tuple[int, int]]:
    """
    Downsample high resolution images to prevent memory issues.

    Returns:
        (processed_img_bytes, was_downsampled, original_size, new_size)
    """
    try:
        img = Image.open(io.BytesIO(img_bytes))
        original_size = img.size  # (width, height)

        max_orig_dimension = max(original_size)

        if max_orig_dimension <= max_dimension:
            return img_bytes, False, original_size, original_size

        scale = max_dimension / max_orig_dimension
        new_size = (int(original_size[0] * scale), int(original_size[1] * scale))

        # Ensure the short side doesn't collapse below a readable minimum.
        # Very wide title blocks (e.g. 26194×4053) would shrink to 4096×633
        # which loses text. Keep short side >= 1500px when possible, but
        # NEVER let the long side exceed max_dimension — that causes 400
        # errors from gpt-5.4 for panoramic RAFAEL title blocks.
        min_short_side = 1500
        short_side = min(new_size)
        if short_side < min_short_side and min(original_size) >= min_short_side:
            up_scale = min_short_side / short_side
            # Cap so the long side stays within max_dimension
            boosted_long = max(new_size) * up_scale
            if boosted_long > max_dimension:
                up_scale = max_dimension / max(new_size)
            new_size = (int(new_size[0] * up_scale), int(new_size[1] * up_scale))

        img_resized = img.resize(new_size, Image.Resampling.LANCZOS)

        output = io.BytesIO()
        img_resized.save(output, format="JPEG", quality=95)
        return output.getvalue(), True, original_size, new_size

    except Exception as e:
        print(f"      Error downsampling image: {e}")
        return img_bytes, False, (0, 0), (0, 0)


def _enhance_contrast_for_title_block(
    img_bytes: bytes,
) -> Tuple[bytes, bool, Dict[str, float]]:
    """
    שיפור ניגודיות עבור Title Block extraction

    פותר בעיית "בהירות יתר" ב-100+ שרטוטים
    """
    try:
        nparr = np.frombuffer(img_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

        if img is None:
            return img_bytes, False, {}

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        original_brightness = float(np.mean(gray))
        original_contrast = float(np.std(gray))

        needs_enhancement = original_brightness > 240 or original_contrast < 35

        if not needs_enhancement:
            return img_bytes, False, {
                "original_brightness": original_brightness,
                "original_contrast": original_contrast,
            }

        print(
            f"     Enhancing contrast (brightness: {original_brightness:.0f}, "
            f"contrast: {original_contrast:.0f})..."
        )

        if original_brightness > 240:
            clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
            enhanced = clahe.apply(gray)
            enhanced = cv2.convertScaleAbs(enhanced, alpha=1.2, beta=-40)
        else:
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(gray)

        kernel = np.array([[-1, -1, -1], [-1, 10, -1], [-1, -1, -1]]) / 2
        enhanced = cv2.filter2D(enhanced, -1, kernel)

        enhanced_brightness = float(np.mean(enhanced))
        enhanced_contrast = float(np.std(enhanced))

        if enhanced_brightness < original_brightness or enhanced_contrast > original_contrast:
            _, buffer = cv2.imencode(".png", enhanced)
            enhanced_bytes = buffer.tobytes()

            print(
                f"     Enhanced: brightness {original_brightness:.0f}→{enhanced_brightness:.0f}, "
                f"contrast {original_contrast:.0f}→{enhanced_contrast:.0f}"
            )

            return enhanced_bytes, True, {
                "original_brightness": original_brightness,
                "enhanced_brightness": enhanced_brightness,
                "original_contrast": original_contrast,
                "enhanced_contrast": enhanced_contrast,
            }
        else:
            return img_bytes, False, {
                "original_brightness": original_brightness,
                "original_contrast": original_contrast,
            }

    except Exception as e:
        print(f"      Contrast enhancement failed: {e}")
        return img_bytes, False, {}


def _extract_image_smart(
    file_path: Path,
    page_num: int = 1,
    dpi: int = 200,
) -> Tuple[bytes, Optional[bytes], Dict[str, bytes], bool]:
    """
    Smart image extraction with high-resolution region support.

    For high-res drawings (>3000px), extracts:
    - Overview at lower DPI (for general extraction)
    - Title Block at high DPI with enlargement (for precise field extraction)

    Returns:
        (overview_bytes, title_block_bytes, additional_regions, used_high_res)
    """
    import fitz

    try:
        with fitz.open(file_path) as doc:
            if page_num > len(doc):
                raise ValueError(f"Page {page_num} does not exist")

            page = doc[page_num - 1]

            initial_pix = page.get_pixmap(dpi=72)
            base_width, base_height = initial_pix.width, initial_pix.height

            scale = dpi / 72
            estimated_width = int(base_width * scale)
            estimated_height = int(base_height * scale)
            max_dimension = max(estimated_width, estimated_height)

            print(f"     Estimated dimensions at DPI {dpi}: {estimated_width}x{estimated_height}px")

            if max_dimension > 3000:
                print(f"     High-res drawing detected ({max_dimension}px) - using region extraction")

                overview_dpi = 200
                print(f"     Creating overview at DPI {overview_dpi}...")
                overview_pix = page.get_pixmap(dpi=overview_dpi)
                overview_bytes = overview_pix.tobytes("png")
                print(f"     Overview: {overview_pix.width}x{overview_pix.height}px")

                # Dynamic DPI — huge drawings don't need 400 DPI
                # At 200 DPI, 18K px drawings already have plenty of detail
                if max_dimension > 12000:
                    highres_dpi = dpi  # Keep original DPI (200) — already huge
                    print(f"     Very large drawing ({max_dimension}px) — using DPI {highres_dpi} (no upscale)")
                elif max_dimension > 8000:
                    highres_dpi = 300  # Medium-large — slight upscale
                    print(f"     Large drawing ({max_dimension}px) — using DPI {highres_dpi}")
                else:
                    highres_dpi = 400  # Normal drawings — full upscale
                    print(f"     Extracting high-res regions at DPI {highres_dpi}...")

                highres_pix = page.get_pixmap(dpi=highres_dpi)
                highres_width, highres_height = highres_pix.width, highres_pix.height
                print(f"     High-res full: {highres_width}x{highres_height}px")

                img_data = highres_pix.samples
                img_np = np.frombuffer(img_data, dtype=np.uint8).reshape(
                    highres_height, highres_width, highres_pix.n
                )

                if highres_pix.n == 4:
                    img_np = cv2.cvtColor(img_np, cv2.COLOR_CMYK2RGB)
                elif highres_pix.n == 3:
                    img_np = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)

                tb_height = int(highres_height * 0.35)
                tb_width = int(highres_width * 0.40)
                tb_region = img_np[highres_height - tb_height:, highres_width - tb_width:]

                # Enlarge factor — reduce for huge drawings
                if max_dimension > 12000:
                    enlarge_factor = 1.0  # No enlargement needed
                elif max_dimension > 8000:
                    enlarge_factor = 1.5
                else:
                    enlarge_factor = 2.0
                tb_enlarged = cv2.resize(
                    tb_region, None, fx=enlarge_factor, fy=enlarge_factor,
                    interpolation=cv2.INTER_CUBIC,
                )

                _, tb_buffer = cv2.imencode(".png", tb_enlarged)
                title_block_bytes = tb_buffer.tobytes()
                print(
                    f"     title_block: {tb_width}x{tb_height} → "
                    f"{tb_enlarged.shape[1]}x{tb_enlarged.shape[0]}px ({enlarge_factor})"
                )

                additional_regions: Dict[str, bytes] = {}

                notes_width = int(highres_width * 0.40)
                notes_height = int(highres_height * 0.35)
                notes_start_y = highres_height - tb_height - notes_height
                if notes_start_y < 0:
                    notes_start_y = 0
                    notes_height = highres_height - tb_height
                notes_region = img_np[
                    notes_start_y : highres_height - tb_height,
                    highres_width - notes_width :,
                ]

                # Dynamic enlarge factor for notes — same logic as title block
                if max_dimension > 12000:
                    notes_enlarge = 1.0
                elif max_dimension > 8000:
                    notes_enlarge = 1.2
                else:
                    notes_enlarge = 1.8
                notes_enlarged = cv2.resize(
                    notes_region, None, fx=notes_enlarge, fy=notes_enlarge, interpolation=cv2.INTER_CUBIC
                )
                # Safety: cap at 30M pixels (Tesseract limit ~32M)
                notes_pixels = notes_enlarged.shape[0] * notes_enlarged.shape[1]
                if notes_pixels > 30_000_000:
                    scale_down = (30_000_000 / notes_pixels) ** 0.5
                    notes_enlarged = cv2.resize(notes_enlarged, None, fx=scale_down, fy=scale_down, interpolation=cv2.INTER_AREA)
                _, notes_buffer = cv2.imencode(".png", notes_enlarged)
                additional_regions["notes"] = notes_buffer.tobytes()
                print(
                    f"     notes: {notes_width}x{notes_height} → "
                    f"{notes_enlarged.shape[1]}x{notes_enlarged.shape[0]}px ({notes_enlarge})"
                )

                mat_width = int(highres_width * 0.25)
                mat_height = int(highres_height * 0.15)
                mat_region = img_np[highres_height - mat_height:, :mat_width]
                # Dynamic enlarge factor for material box
                if max_dimension > 12000:
                    mat_enlarge = 1.0
                elif max_dimension > 8000:
                    mat_enlarge = 1.2
                else:
                    mat_enlarge = 1.8
                mat_enlarged = cv2.resize(
                    mat_region, None, fx=mat_enlarge, fy=mat_enlarge, interpolation=cv2.INTER_CUBIC
                )
                # Safety: cap at 30M pixels
                mat_pixels = mat_enlarged.shape[0] * mat_enlarged.shape[1]
                if mat_pixels > 30_000_000:
                    scale_down = (30_000_000 / mat_pixels) ** 0.5
                    mat_enlarged = cv2.resize(mat_enlarged, None, fx=scale_down, fy=scale_down, interpolation=cv2.INTER_AREA)
                _, mat_buffer = cv2.imencode(".png", mat_enlarged)
                additional_regions["material_box"] = mat_buffer.tobytes()
                print(
                    f"     material_box: {mat_width}x{mat_height} → "
                    f"{mat_enlarged.shape[1]}x{mat_enlarged.shape[0]}px ({mat_enlarge})"
                )

                return overview_bytes, title_block_bytes, additional_regions, True

            else:
                print("     Standard resolution drawing - using single extraction")
                pix = page.get_pixmap(dpi=dpi)
                img_bytes_out = pix.tobytes("png")
                print(f"     Extracted: {pix.width}x{pix.height}px at DPI {dpi}")
                return img_bytes_out, None, {}, False

    except Exception as e:
        print(f"     Error in smart extraction: {e}")
        try:
            with fitz.open(file_path) as doc:
                page = doc[page_num - 1]
                pix = page.get_pixmap(dpi=dpi)
                img_bytes_out = pix.tobytes("png")
                return img_bytes_out, None, {}, False
        except Exception as e2:
            print(f"     Fallback extraction also failed: {e2}")
            raise


def _assess_image_quality(img_bytes: bytes) -> Dict[str, Any]:
    """
    Assess image quality for OCR readiness.

    Checks sharpness, contrast, brightness.
    """
    quality_info: Dict[str, Any] = {
        "sharpness_score": 0,
        "contrast_score": 0,
        "brightness_score": 0,
        "is_blurry": False,
        "is_low_contrast": False,
        "is_too_dark": False,
        "is_too_bright": False,
        "quality_issues": [],
    }

    try:
        nparr = np.frombuffer(img_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

        if img is None:
            return quality_info

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        laplacian_var = cv2.Laplacian(gray, cv2.CV_64F).var()
        quality_info["sharpness_score"] = round(laplacian_var, 2)

        if laplacian_var < 100:
            quality_info["is_blurry"] = True
            if laplacian_var < 50:
                quality_info["quality_issues"].append(
                    f"מטושטש מאוד (חדות: {laplacian_var:.0f})"
                )
            else:
                quality_info["quality_issues"].append(
                    f"מטושטש (חדות: {laplacian_var:.0f})"
                )

        contrast = gray.std()
        quality_info["contrast_score"] = round(contrast, 2)
        if contrast < 30:
            quality_info["is_low_contrast"] = True
            quality_info["quality_issues"].append(f"ניגודיות נמוכה ({contrast:.0f})")

        brightness = gray.mean()
        quality_info["brightness_score"] = round(brightness, 2)
        if brightness < 60:
            quality_info["is_too_dark"] = True
            quality_info["quality_issues"].append(
                f"כהה מדי (בהירות: {brightness:.0f})"
            )
        elif brightness > 200:
            quality_info["is_too_bright"] = True
            quality_info["quality_issues"].append(
                f"בהיר מדי (בהירות: {brightness:.0f})"
            )

    except Exception as e:
        quality_info["quality_issues"].append(f"שגיאה בבדיקת איכות: {str(e)}")

    return quality_info


def _validate_rotation_improvement(
    original_bytes: bytes,
    rotated_bytes: bytes,
) -> Tuple[bool, float]:
    """
    Validate if rotation actually improved OCR clarity.

    Returns:
        (rotation_helped, improvement_ratio)
    """
    try:
        import pytesseract

        nparr_orig = np.frombuffer(original_bytes, np.uint8)
        img_orig = cv2.imdecode(nparr_orig, cv2.IMREAD_COLOR)

        nparr_rot = np.frombuffer(rotated_bytes, np.uint8)
        img_rot = cv2.imdecode(nparr_rot, cv2.IMREAD_COLOR)

        if img_orig is None or img_rot is None:
            return False, 1.0

        gray_orig = cv2.cvtColor(img_orig, cv2.COLOR_BGR2GRAY)
        gray_rot = cv2.cvtColor(img_rot, cv2.COLOR_BGR2GRAY)

        h_orig, w_orig = gray_orig.shape[:2]
        crop_size = min(300, h_orig // 3, w_orig // 3)
        crop_y = (h_orig - crop_size) // 2
        crop_x = (w_orig - crop_size) // 2
        sample_orig = gray_orig[crop_y : crop_y + crop_size, crop_x : crop_x + crop_size]

        h_rot, w_rot = gray_rot.shape[:2]
        crop_y_rot = (h_rot - crop_size) // 2
        crop_x_rot = (w_rot - crop_size) // 2
        sample_rot = gray_rot[
            crop_y_rot : crop_y_rot + crop_size, crop_x_rot : crop_x_rot + crop_size
        ]

        try:
            text_orig = pytesseract.image_to_string(sample_orig)
            score_orig = len([c for c in text_orig if c.isalnum()])
        except Exception as e:
            debug_print(f"[ROTATION_VALIDATION_ORIG_OCR] Error: {e}")
            score_orig = 0

        try:
            text_rot = pytesseract.image_to_string(sample_rot)
            score_rot = len([c for c in text_rot if c.isalnum()])
        except Exception as e:
            debug_print(f"[ROTATION_VALIDATION_ROT_OCR] Error: {e}")
            score_rot = 0

        if score_orig > 0:
            improvement_ratio = score_rot / score_orig
        else:
            improvement_ratio = 1.0

        rotation_helped = improvement_ratio > 1.1

        print(
            f"        ✓ Rotation validation: Original {score_orig} chars → "
            f"Rotated {score_rot} chars (ratio: {improvement_ratio:.2f})"
        )

        if not rotation_helped:
            print(
                f"        ⚠ Rotation did NOT improve OCR "
                f"(only {improvement_ratio:.1%}), reverting..."
            )

        return rotation_helped, improvement_ratio

    except Exception as e:
        print(f"        ⚠ Validation failed: {e}, cannot confirm improvement")
        return False, 1.0


def _apply_rotation_angle(img_bytes: bytes, rotation_angle: int) -> bytes:
    """
    Apply a known rotation angle to image bytes WITHOUT checking rotation.
    Used for regions extracted from the same page (they share the same rotation).
    """
    if rotation_angle == 0:
        return img_bytes

    try:
        img = Image.open(io.BytesIO(img_bytes))
        img_rotated = img.rotate(rotation_angle, expand=True, fillcolor="white")
        output = io.BytesIO()
        img_rotated.save(output, format="JPEG", quality=95)
        return output.getvalue()
    except Exception as e:
        debug_print(f"[APPLY_ROTATION_ANGLE] Error: {e}")
        return img_bytes


def _estimate_quarter_turn_hint(img_bytes: bytes) -> Optional[int]:
    """
    Heuristic hint for drawings that are sideways (90°/270°) rather than 180°.
    Returns CCW rotation suggestion: +90, -90, or 0 (none).
    """
    try:
        nparr = np.frombuffer(img_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        if img is None:
            return 0

        h, w = img.shape[:2]
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, bw = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
        bw = cv2.morphologyEx(bw, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))

        band = max(40, int(min(h, w) * 0.08))
        left = int(np.count_nonzero(bw[:, :band]))
        right = int(np.count_nonzero(bw[:, w - band :]))
        top = int(np.count_nonzero(bw[:band, :]))
        bottom = int(np.count_nonzero(bw[h - band :, :]))

        lr = max(left, right)
        tb = max(top, bottom)

        if lr > tb * 1.25:
            return 90 if left >= right else -90

        return 0
    except Exception:
        return 0


def _fix_image_rotation(
    img_bytes: bytes,
    file_path: Optional[str] = None,
    skip_azure_check: bool = False,
) -> Tuple[bytes, bool, str]:
    """
    Detect and correct image rotation.

    Attempts:
    1. Check EXIF data first (fastest method)
    2. If no EXIF and skip_azure_check=False, use Azure OpenAI Vision
    3. Fallback to OCR confidence method
    """
    # Lazy imports that still live in the vision_api module
    from src.services.ai.vision_api import (
        _build_client,
        _chat_create_with_token_compat,
        _resolve_stage_call_config,
    )

    # CHECK CACHE FIRST
    if file_path and file_path in _rotation_cache:
        was_rotated, rotation_angle, rotation_info = _rotation_cache[file_path]
        if rotation_angle != 0:
            try:
                img = Image.open(io.BytesIO(img_bytes))
                img_rotated = img.rotate(rotation_angle, expand=True, fillcolor="white")
                output = io.BytesIO()
                img_rotated.save(output, format="JPEG", quality=95)
                return output.getvalue(), was_rotated, rotation_info
            except Exception as e:
                debug_print(f"[ROTATION_CACHE_APPLY] Error: {e}")
                pass  # Proceed with full check if cache application fails
        else:
            return img_bytes, was_rotated, rotation_info

    rotation_info = "✓ No rotation detected"
    try:
        img = Image.open(io.BytesIO(img_bytes))

        # Step 1: Check EXIF orientation data
        try:
            exif_data = img._getexif()
            if exif_data:
                for tag_id, value in exif_data.items():
                    tag_name = TAGS.get(tag_id, tag_id)
                    if tag_name == "Orientation":
                        if value == 6:
                            img = img.rotate(270, expand=True)
                            rotation_info = "⚠ Rotated 90° CW (EXIF)"
                        elif value == 8:
                            img = img.rotate(90, expand=True)
                            rotation_info = "⚠ Rotated 270° CW (EXIF)"
                        elif value == 3:
                            img = img.rotate(180, expand=True)
                            rotation_info = "⚠ Rotated 180° (EXIF)"
                        elif value in [2, 4, 5, 7]:
                            if value in [2, 4]:
                                img = img.transpose(Image.Transpose.FLIP_LEFT_RIGHT)
                            if value in [5, 7]:
                                img = img.transpose(Image.Transpose.FLIP_TOP_BOTTOM)
                            rotation_info = f"⚠ Flipped/Mirrored (EXIF #{value})"

                        if rotation_info != "No rotation detected":
                            output = io.BytesIO()
                            img.save(output, format="JPEG", quality=95)
                            if file_path:
                                _rotation_cache[file_path] = (
                                    True,
                                    90 if "90" in rotation_info else 180 if "180" in rotation_info else 270,
                                    rotation_info,
                                )
                            return output.getvalue(), True, rotation_info
        except (AttributeError, KeyError, IndexError):
            pass

        # Pre-hint: detect likely quarter-turn pages (sideways drawings)
        quarter_turn_hint = _estimate_quarter_turn_hint(img_bytes)
        if quarter_turn_hint in (90, -90):
            print(f"        📐 Orientation hint: likely sideways page (suggest {quarter_turn_hint}° CCW)")

        # Step 2a: Azure OpenAI Vision rotation detection
        if not skip_azure_check:
            try:
                print("        📐 Asking Azure Vision about rotation...")
                client = _build_client()

                img_b64_check = base64.b64encode(img_bytes).decode("utf-8")

                rotation_prompt = load_prompt("10_detect_rotation")

                rotation_model, rotation_max_tokens, rotation_temperature = _resolve_stage_call_config(
                    STAGE_ROTATION, 10, 0,
                )

                response = _chat_create_with_token_compat(
                    client,
                    model=rotation_model,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": rotation_prompt},
                                {
                                    "type": "image_url",
                                    "image_url": {"url": f"data:image/jpeg;base64,{img_b64_check}"},
                                },
                            ],
                        }
                    ],
                    max_tokens=rotation_max_tokens,
                    temperature=rotation_temperature,
                )

                rotation_answer = response.choices[0].message.content.strip().lower()
                print(f"        📐 Azure Vision response: '{rotation_answer}'")

                try:
                    numbers = re.findall(r"\d+", rotation_answer)
                    if numbers:
                        detected_angle = int(numbers[0])

                        if detected_angle > 180:
                            detected_angle = detected_angle - 360

                        rotation_angle = -detected_angle

                        if abs(abs(detected_angle) - 180) <= 10 and quarter_turn_hint in (90, -90):
                            print(
                                f"        ⚠ Azure suggested ~180°, but geometry indicates "
                                f"sideways; using {quarter_turn_hint}° hint"
                            )
                            rotation_angle = quarter_turn_hint

                        if abs(rotation_angle) > 2:
                            img_rotated = img.rotate(rotation_angle, expand=True, fillcolor="white")

                            output = io.BytesIO()
                            img_rotated.save(output, format="JPEG", quality=95)
                            rotated_img_bytes = output.getvalue()

                            rotation_helped, improvement_ratio = _validate_rotation_improvement(
                                img_bytes, rotated_img_bytes
                            )

                            accept_by_hint = (
                                (not rotation_helped)
                                and quarter_turn_hint in (90, -90)
                                and rotation_angle == quarter_turn_hint
                                and abs(rotation_angle) == 90
                                and improvement_ratio >= 0.95
                            )

                            if rotation_helped or accept_by_hint:
                                rotation_info = (
                                    f"⚠ Rotated {rotation_angle}° CCW "
                                    f"(Azure detected {detected_angle}°, validation: {improvement_ratio:.1%})"
                                )
                                print(
                                    f"        ✅ ROTATION FIXED BY AZURE VISION! "
                                    f"(precise: {detected_angle}°, validated)"
                                )
                                if file_path:
                                    _rotation_cache[file_path] = (True, rotation_angle, rotation_info)
                                return rotated_img_bytes, True, rotation_info
                            else:
                                print(
                                    f"        ⚠ Rotation REJECTED by validation "
                                    f"(score degraded to {improvement_ratio:.1%}), keeping original"
                                )
                                rotation_info = (
                                    f"✓ Rotation rejected by validation "
                                    f"(Azure said {detected_angle}° but OCR quality declined)"
                                )
                                if file_path:
                                    _rotation_cache[file_path] = (False, 0, rotation_info)
                                return img_bytes, False, rotation_info
                        else:
                            print("        ✓ Azure: Rotation < 2° threshold")
                            rotation_info = "✓ Azure Vision: No significant rotation"
                            if file_path:
                                _rotation_cache[file_path] = (False, 0, rotation_info)
                            return img_bytes, False, rotation_info
                    else:
                        if "not rotated" in rotation_answer or rotation_answer == "0":
                            print("        ✓ Azure: No rotation detected")
                            rotation_info = "✓ Azure Vision: Not rotated"
                            if file_path:
                                _rotation_cache[file_path] = (False, 0, rotation_info)
                            return img_bytes, False, rotation_info
                        else:
                            print(
                                f"        ⚠ Azure response unclear: {rotation_answer}, "
                                f"trying fallback..."
                            )
                except (ValueError, IndexError):
                    print(
                        f"        ⚠ Could not parse Azure response: {rotation_answer}, "
                        f"trying fallback..."
                    )

            except Exception as e:
                print(f"        ⚠ Azure Vision rotation detection failed: {e}")
            else:
                if rotation_info == "✓ Azure Vision: Not rotated":
                    return img_bytes, False, rotation_info

        # Step 2b: Fallback to OpenCV/OCR
        nparr = np.frombuffer(img_bytes, np.uint8)
        img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

        if img_cv is None:
            return img_bytes, False, rotation_info

        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

        print("        📐 Testing rotations [OCR confidence method - FINER GRAIN]...")

        best_rotation = 0
        best_ocr_score = 0

        test_angles = list(range(0, 360, 5))
        ocr_results: Dict[int, int] = {}

        try:
            h, w = gray.shape[:2]
            crop_size = min(400, h // 2, w // 2)
            crop_y = (h - crop_size) // 2
            crop_x = (w - crop_size) // 2
            gray_crop = gray[crop_y : crop_y + crop_size, crop_x : crop_x + crop_size]

            for angle in test_angles:
                mat = cv2.getRotationMatrix2D((crop_size // 2, crop_size // 2), angle, 1.0)
                rotated = cv2.warpAffine(
                    gray_crop, mat, (crop_size, crop_size),
                    borderMode=cv2.BORDER_CONSTANT, borderValue=255,
                )

                try:
                    import pytesseract

                    text = pytesseract.image_to_string(rotated)
                    score = len([c for c in text if c.isalnum()])
                    ocr_results[angle] = score

                    if score > best_ocr_score:
                        best_ocr_score = score
                        best_rotation = angle
                except Exception:
                    ocr_results[angle] = 0

            if best_ocr_score > 10:
                print(
                    f"        📐 OCR confidence scores (top 5): "
                    f"{sorted(ocr_results.items(), key=lambda x: x[1], reverse=True)[:5]}"
                )
                print(f"        📐 Best rotation: {best_rotation}° (score: {best_ocr_score})")

                if best_rotation > 180:
                    best_rotation = best_rotation - 360

                if abs(best_rotation) > 2:
                    h, w = img_cv.shape[:2]
                    center = (w // 2, h // 2)
                    mat = cv2.getRotationMatrix2D(center, best_rotation, 1.0)
                    img_rotated = cv2.warpAffine(
                        img_cv, mat, (w, h),
                        borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255),
                    )

                    _, buf = cv2.imencode(".jpg", img_rotated)
                    rotation_info = (
                        f"⚠ Rotated {best_rotation}° CCW "
                        f"(OCR confidence detection - fine grain)"
                    )
                    print("        ✅ ROTATION FIXED!")
                    if file_path:
                        _rotation_cache[file_path] = (True, best_rotation, rotation_info)
                    return buf.tobytes(), True, rotation_info
        except Exception as e:
            print(f"        📐 OCR method failed: {e}, trying fallback...")

        # FALLBACK: Line angle detection
        edges = cv2.Canny(gray, 50, 150)
        lines = cv2.HoughLines(edges, 1, np.pi / 180, 30)

        if lines is not None and len(lines) > 10:
            text_angles = []
            for line in lines:
                rho, theta = line[0]
                angle = np.degrees(theta)
                normalized = angle if angle <= 90 else angle - 180
                if not (
                    abs(normalized) < 10
                    or abs(normalized - 90) < 10
                    or abs(normalized + 90) < 10
                ):
                    text_angles.append(normalized)

            if text_angles:
                median_text_angle = float(np.median(text_angles))

                if abs(median_text_angle) > 10:
                    h, w = img_cv.shape[:2]
                    center = (w // 2, h // 2)
                    mat = cv2.getRotationMatrix2D(center, median_text_angle, 1.0)
                    img_rotated = cv2.warpAffine(
                        img_cv, mat, (w, h),
                        borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255),
                    )

                    _, buf = cv2.imencode(".jpg", img_rotated)
                    rotation_info = f"⚠ Rotated {median_text_angle:.1f}° CCW (Line angle fallback)"
                    print("        ✅ ROTATION FIXED (fallback)!")
                    return buf.tobytes(), True, rotation_info

        rotation_info = "✓ Unable to detect significant rotation"
        return img_bytes, False, rotation_info

    except Exception as e:
        print(f"      ⚠ Error in rotation detection: {e}")
        return img_bytes, False, f"Error in rotation detection: {str(e)}"
