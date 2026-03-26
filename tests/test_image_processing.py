"""Tests for image/processing.py (no API needed)."""
import io
import pytest
import numpy as np

# Create test image fixture
@pytest.fixture
def small_image_bytes():
    """Create a small 200x100 test PNG image."""
    from PIL import Image
    img = Image.new("RGB", (200, 100), color=(128, 128, 128))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

@pytest.fixture
def large_image_bytes():
    """Create a 6000x4000 test PNG image (above max_dim=4096)."""
    from PIL import Image
    img = Image.new("RGB", (6000, 4000), color=(200, 200, 200))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


class TestDownsample:
    def test_small_image_unchanged(self, small_image_bytes):
        from src.services.image.processing import _downsample_high_res_image
        result_bytes, was_downsampled, orig_size, new_size = _downsample_high_res_image(small_image_bytes, max_dimension=4096)
        # Small image should pass through (or be re-encoded at same size)
        assert len(result_bytes) > 0
        assert not was_downsampled

    def test_large_image_reduced(self, large_image_bytes):
        from src.services.image.processing import _downsample_high_res_image
        result_bytes, was_downsampled, orig_size, new_size = _downsample_high_res_image(large_image_bytes, max_dimension=4096)
        from PIL import Image
        img = Image.open(io.BytesIO(result_bytes))
        assert max(img.size) <= 4096
        assert was_downsampled

    def test_returns_tuple(self, small_image_bytes):
        from src.services.image.processing import _downsample_high_res_image
        result = _downsample_high_res_image(small_image_bytes)
        assert isinstance(result, tuple)
        assert len(result) == 4
        assert isinstance(result[0], bytes)


class TestImageQuality:
    def test_assess_returns_dict(self, small_image_bytes):
        from src.services.image.processing import _assess_image_quality
        result = _assess_image_quality(small_image_bytes)
        assert isinstance(result, dict)

    def test_assess_has_quality_fields(self, small_image_bytes):
        from src.services.image.processing import _assess_image_quality
        result = _assess_image_quality(small_image_bytes)
        # Should contain at least some quality metrics
        assert len(result) > 0


class TestApplyRotation:
    def test_rotate_0_returns_same_size(self, small_image_bytes):
        from src.services.image.processing import _apply_rotation_angle
        result = _apply_rotation_angle(small_image_bytes, 0)
        from PIL import Image
        original = Image.open(io.BytesIO(small_image_bytes))
        rotated = Image.open(io.BytesIO(result))
        assert original.size == rotated.size

    def test_rotate_90_swaps_dimensions(self, small_image_bytes):
        from src.services.image.processing import _apply_rotation_angle
        result = _apply_rotation_angle(small_image_bytes, 90)
        from PIL import Image
        original = Image.open(io.BytesIO(small_image_bytes))
        rotated = Image.open(io.BytesIO(result))
        # 200x100 rotated 90° → 100x200
        assert rotated.size[0] == original.size[1]
        assert rotated.size[1] == original.size[0]

    def test_rotate_180_same_dimensions(self, small_image_bytes):
        from src.services.image.processing import _apply_rotation_angle
        result = _apply_rotation_angle(small_image_bytes, 180)
        from PIL import Image
        original = Image.open(io.BytesIO(small_image_bytes))
        rotated = Image.open(io.BytesIO(result))
        assert original.size == rotated.size

    def test_returns_bytes(self, small_image_bytes):
        from src.services.image.processing import _apply_rotation_angle
        result = _apply_rotation_angle(small_image_bytes, 90)
        assert isinstance(result, bytes)


class TestEnhanceContrast:
    def test_returns_tuple(self, small_image_bytes):
        from src.services.image.processing import _enhance_contrast_for_title_block
        result = _enhance_contrast_for_title_block(small_image_bytes)
        assert isinstance(result, tuple)
        assert len(result) == 3
        assert isinstance(result[0], bytes)

    def test_output_is_valid_image(self, small_image_bytes):
        from src.services.image.processing import _enhance_contrast_for_title_block
        from PIL import Image
        result_bytes, was_enhanced, metrics = _enhance_contrast_for_title_block(small_image_bytes)
        img = Image.open(io.BytesIO(result_bytes))
        assert img.size[0] > 0 and img.size[1] > 0
