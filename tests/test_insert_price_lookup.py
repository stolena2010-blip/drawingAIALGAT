"""Tests for src.services.extraction.insert_price_lookup."""

import pytest
from unittest.mock import patch
from src.services.extraction.insert_price_lookup import (
    lookup_insert_price,
    enrich_inserts_with_prices,
    _load_catalog,
    _by_pn,
    _by_name,
    _pn_purchases,
)


@pytest.fixture(autouse=True)
def _ensure_catalog_loaded():
    """Make sure the catalog is loaded (uses real BOM/INSERTS.xlsx)."""
    _load_catalog()


class TestLookupInsertPrice:
    """Test price lookup against the real catalog."""

    def test_rafael_402_pn(self):
        """402xxxxxx direct PN match."""
        result = lookup_insert_price("402000017")
        assert result is not None
        price, currency = result
        assert price > 0
        assert currency in ('₪', '$')

    def test_ms_standard_in_name(self):
        """MS51830 found via name token."""
        result = lookup_insert_price("MS51830-201")
        assert result is not None

    def test_ma3279_match(self):
        """MA3279-xxx direct match."""
        result = lookup_insert_price("MA3279-154")
        assert result is not None
        price, _ = result
        assert price > 0

    def test_unknown_pn_returns_none(self):
        """Unknown PN returns None."""
        result = lookup_insert_price("BOGUS999999")
        assert result is None

    def test_empty_pn(self):
        result = lookup_insert_price("")
        assert result is None

    def test_description_fallback(self):
        """Find by description when cat_no doesn't match."""
        result = lookup_insert_price("UNKNOWN", "INSERT HELICAL MS122078")
        assert result is not None

    def test_bp_part_no_price(self):
        """Sub-assembly parts should NOT have prices."""
        result = lookup_insert_price("BP41596A")
        assert result is None

    def test_duplicate_pn_uses_most_purchases(self):
        """When multiple rows share a PN, the one with most purchases wins."""
        _load_catalog()
        # 402050089 has 5 rows; the one with most purchases (14) is 6.7 ₪
        result = lookup_insert_price("402050089")
        assert result is not None
        price, currency = result
        # The most-purchased row has 14 purchases, price 6.7 ₪
        assert price == 6.7
        assert currency == '₪'

    def test_zero_price_rows_skipped(self):
        """Rows with price=0 should be excluded from the catalog."""
        _load_catalog()
        # Confirm that no entry in the catalog has price <= 0
        for pn, (price, _) in _by_pn.items():
            assert price > 0, f"PN {pn} has non-positive price {price}"

    def test_purchase_counts_stored(self):
        """Purchase counts are tracked for deduplication."""
        _load_catalog()
        # 402050089 should have purchase count = 14 (the winning row)
        assert _pn_purchases.get("402050089", 0) == 14


class TestEnrichInsertsWithPrices:
    """Test the enrichment function."""

    def test_adds_price_fields(self):
        items = [
            {"cat_no": "402000017", "qty": "4", "description": "INSERT HELICAL"},
        ]
        result = enrich_inserts_with_prices(items)
        assert result[0].get("unit_price") is not None
        assert result[0]["unit_price"] > 0
        assert result[0]["currency"] in ('₪', '$')

    def test_unknown_gets_none_price(self):
        items = [
            {"cat_no": "BOGUS999", "qty": "2", "description": "SOMETHING"},
        ]
        result = enrich_inserts_with_prices(items)
        assert result[0]["unit_price"] is None
        assert result[0]["currency"] == ''

    def test_empty_list(self):
        assert enrich_inserts_with_prices([]) == []

    def test_none_input(self):
        assert enrich_inserts_with_prices(None) is None

    def test_mixed_items(self):
        items = [
            {"cat_no": "402000017", "qty": "4", "description": "INSERT"},
            {"cat_no": "BOGUS", "qty": "1", "description": "UNKNOWN"},
        ]
        result = enrich_inserts_with_prices(items)
        assert result[0]["unit_price"] is not None  # found
        assert result[1]["unit_price"] is None       # not found
