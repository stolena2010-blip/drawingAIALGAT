"""
Insert Price Lookup — loads prices from BOM/INSERTS.xlsx catalog.

Provides a lookup function that, given a cat_no or description,
returns (unit_price, currency) if found.
"""

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

_INSERTS_PATH = Path(__file__).resolve().parents[3] / "BOM" / "INSERTS.xlsx"

# Cached lookup tables (loaded once)
_by_pn: Dict[str, Tuple[float, str]] = {}
_by_name: Dict[str, Tuple[float, str]] = {}
_pn_purchases: Dict[str, int] = {}      # purchase count per PN (for dedup)
_name_purchases: Dict[str, int] = {}    # purchase count per name token
_loaded = False


def _clean_pn(raw: str) -> str:
    """Normalize a raw PN from the catalog for lookup."""
    s = raw.strip()
    # Remove prefixes: "RAF PN: ", "רפאל ", "Rafael ", " רפאל "
    s = re.sub(r'^(RAF\s*PN\s*:?\s*|רפאל\s+|Rafael\s+)', '', s, flags=re.IGNORECASE).strip()
    return s.upper()


def _load_catalog() -> None:
    """Load BOM/INSERTS.xlsx into memory once.

    When multiple rows share the same PN, the row with the highest
    כמות רכישות (purchase count) wins.  Rows with price ≤ 0 or
    invalid price are skipped.
    """
    global _by_pn, _by_name, _pn_purchases, _name_purchases, _loaded
    if _loaded:
        return

    _loaded = True  # mark even if fails, so we don't retry

    if not _INSERTS_PATH.exists():
        logger.warning(f"[INSERT-PRICE] Catalog not found: {_INSERTS_PATH}")
        return

    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(_INSERTS_PATH), data_only=True, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        logger.error(f"[INSERT-PRICE] Failed to load catalog: {e}")
        return

    if len(rows) < 2:
        return

    # Columns: 0=מק"ט, 1=שם פריט, 2=עלות, 3=מטבע, 4=תאריך, 5=כמות רכישות, ...
    for row in rows[1:]:
        pn_raw = str(row[0] or '').strip()
        name = str(row[1] or '').strip()
        price_val = row[2]
        currency = str(row[3] or '').strip()
        purchases_val = row[5] if len(row) > 5 else 0

        if price_val is None:
            continue
        try:
            price = float(price_val)
        except (ValueError, TypeError):
            continue

        # Skip zero / negative prices
        if price <= 0:
            continue

        if not currency or currency == '#N/A':
            currency = '₪'

        try:
            purchases = int(purchases_val) if purchases_val else 0
        except (ValueError, TypeError):
            purchases = 0

        val = (price, currency)

        # Index by all PN variants in column A
        # Handle multi-PN: "402260086/MA3279-154"
        clean = _clean_pn(pn_raw)
        for pn_part in re.split(r'[/,]', clean):
            pn_part = pn_part.strip()
            if not pn_part:
                continue
            # Keep the row with the most purchases
            if pn_part not in _by_pn or purchases > _pn_purchases.get(pn_part, 0):
                _by_pn[pn_part] = val
                _pn_purchases[pn_part] = purchases

        # Index by standard PN tokens in the name (MS, MA, NAS, KN, IOI)
        for token in re.findall(
            r'(MS\d{5,}[A-Z0-9\-]*|MS\s?\d{5,}[A-Z0-9\-]*|'
            r'MA\d{4}[A-Z0-9\-]*|NAS\d{4}[A-Z0-9\-]*|'
            r'IOI[A-Z0-9\-]+|KN[LMC][A-Z0-9\-]+)',
            name,
            re.IGNORECASE,
        ):
            key = re.sub(r'\s+', '', token).upper()
            if key not in _by_name or purchases > _name_purchases.get(key, 0):
                _by_name[key] = val
                _name_purchases[key] = purchases

    logger.info(f"[INSERT-PRICE] Loaded {len(_by_pn)} PNs + {len(_by_name)} name tokens from catalog")


def lookup_insert_price(cat_no: str, description: str = "") -> Optional[Tuple[float, str]]:
    """
    Look up the price of an insert by its catalog number and/or description.

    Parameters
    ----------
    cat_no : str
        The insert catalog number (e.g. "402050023", "MS21209C0215").
    description : str
        The insert description (e.g. "INSERT HELICAL MS51830").

    Returns
    -------
    (price, currency) or None if not found.
    """
    _load_catalog()

    if not cat_no:
        return None
    cat_upper = cat_no.strip().upper()
    if not cat_upper:
        return None

    # 1. Direct PN match
    if cat_upper in _by_pn:
        return _by_pn[cat_upper]

    # 2. Try cleaning the PN (remove RAF PN: prefix etc.)
    clean = _clean_pn(cat_no)
    if clean in _by_pn:
        return _by_pn[clean]

    # 3. Try matching standard tokens from cat_no
    for token in re.findall(
        r'(MS\d{5,}[A-Z0-9\-]*|MA\d{4}[A-Z0-9\-]*|NAS\d{4}[A-Z0-9\-]*|IOI[A-Z0-9\-]+|KN[LMC][A-Z0-9\-]+)',
        cat_upper,
    ):
        key = re.sub(r'\s+', '', token)
        if key in _by_name:
            return _by_name[key]

    # 4. Try tokens from description
    if description:
        desc_upper = description.strip().upper()
        for token in re.findall(
            r'(MS\d{5,}[A-Z0-9\-]*|MS\s?\d{5,}[A-Z0-9\-]*|'
            r'MA\d{4}[A-Z0-9\-]*|NAS\d{4}[A-Z0-9\-]*|'
            r'IOI[A-Z0-9\-]+|KN[LMC][A-Z0-9\-]+)',
            desc_upper,
        ):
            key = re.sub(r'\s+', '', token)
            if key in _by_name:
                return _by_name[key]

    # 5. Substring match on PN catalog (last resort — for 402xxx in longer strings)
    if len(cat_upper) >= 6:
        for catalog_pn, val in _by_pn.items():
            if len(catalog_pn) >= 6 and (catalog_pn in cat_upper or cat_upper in catalog_pn):
                return val

    return None


def enrich_inserts_with_prices(
    inserts_list: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Add unit_price and currency fields to each insert dict.

    Modifies items in-place and returns the same list.
    """
    if not inserts_list:
        return inserts_list

    for item in inserts_list:
        if not isinstance(item, dict):
            continue
        cat_no = item.get('cat_no', '')
        desc = item.get('description', '')
        result = lookup_insert_price(cat_no, desc)
        if result:
            item['unit_price'] = result[0]
            item['currency'] = result[1]
        else:
            item['unit_price'] = None
            item['currency'] = ''

    return inserts_list
