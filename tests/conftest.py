"""Shared fixtures for DrawingAI Pro tests."""
import json
import sys
from pathlib import Path

import pytest

# Ensure project root is in path
sys.path.insert(0, str(Path(__file__).parent.parent))

FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture
def sample_results():
    """Fake drawing results as returned by extract_customer_name."""
    return [
        {
            "file_name": "drawing_12345.pdf",
            "part_number": "12345-A",
            "drawing_number": "DWG-001",
            "revision": "B",
            "item_name": "BRACKET",
            "material": "AL 6061",
            "processes": "Anodizing, Painting",
            "notes": "MIL-A-8625 Type III",
            "geometric_area": "~300 cm",
            "customer_name": "IAI",
            "confidence_level": "HIGH",
            "email_from": "test@customer.com",
            "email_subject": "Order 5678",
        },
        {
            "file_name": "drawing_67890.pdf",
            "part_number": "67890-C",
            "drawing_number": "DWG-002",
            "revision": "A",
            "item_name": "COVER PLATE",
            "material": "SS 304",
            "processes": "Passivation",
            "notes": "AMS 2700",
            "geometric_area": "~150 cm",
            "customer_name": "RAFAEL",
            "confidence_level": "FULL",
            "email_from": "test@customer.com",
            "email_subject": "Quote 9012",
        },
    ]


@pytest.fixture
def sample_classifications():
    """Fake file classification list."""
    return [
        {
            "file_path": Path("/tmp/test/drawing_12345.pdf"),
            "file_type": "DRAWING",
            "associated_item": "12345-A",
            "item_number": "DWG-001",
            "revision": "B",
            "original_filename": "drawing_12345.pdf",
        },
        {
            "file_path": Path("/tmp/test/drawing_67890.pdf"),
            "file_type": "DRAWING",
            "associated_item": "67890-C",
            "item_number": "DWG-002",
            "revision": "A",
            "original_filename": "drawing_67890.pdf",
        },
        {
            "file_path": Path("/tmp/test/order.pdf"),
            "file_type": "PURCHASE_ORDER",
            "associated_item": "12345-A",
            "item_number": "",
            "revision": "",
            "original_filename": "order.pdf",
        },
    ]


@pytest.fixture
def sample_pl_items():
    """Fake PL extraction results."""
    return [
        {
            "pl_filename": "PL_12345.pdf",
            "item_number": "FASTENER-A1",
            "description": "Hex bolt M6x20",
            "quantity": "12",
            "processes": ["passivation"],
            "specifications": ["AMS 2700"],
            "associated_item": "12345-A",
            "matched_item_name": "12345-A",
        }
    ]


@pytest.fixture
def tmp_output(tmp_path):
    """Temporary output directory."""
    return tmp_path
