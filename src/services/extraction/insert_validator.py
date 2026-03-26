"""
Insert Validator — filters out non-insert BOM items from inserts_hardware.

Uses keyword-based matching derived from the known inserts catalog (BOM/INSERTS.xlsx).
Real inserts match at least one of the known patterns (description or cat_no).
BOM items that are sub-assembly parts/components are filtered out.
"""

import logging
import re
from typing import Any, Dict, List

logger = logging.getLogger(__name__)

# ── Known insert keywords (from BOM/INSERTS.xlsx analysis) ────────
# These appear in the DESCRIPTION of real inserts
_INSERT_DESC_KEYWORDS: set[str] = {
    "INSERT",
    "HELICAL",
    "HELICOIL",
    "KEENSERT",
    "KEY LOCKED",
    "KEY-LOCKED",
    "COIL INSERT",
    "RIVNUT",
    "PEM",
    "SELF CLINCH",
    "STANDOFF",
    "BUSHING",
    "LOCTITE",
    "LOCKTITE",
    "BOLLHOFF",
}

# Standard part number prefixes for real inserts
_INSERT_PN_PATTERNS: list[re.Pattern] = [
    re.compile(r"^MS51830", re.IGNORECASE),   # Key-locked inserts
    re.compile(r"^MS51831", re.IGNORECASE),   # Key-locked inserts (screw-thread)
    re.compile(r"^MS21209", re.IGNORECASE),   # Helical inserts
    re.compile(r"^MS122\d", re.IGNORECASE),   # Helical inserts (MS122xxx)
    re.compile(r"^MS124\d", re.IGNORECASE),   # Helical inserts (MS124xxx)
    re.compile(r"^MA3279", re.IGNORECASE),    # Helicoil standard
    re.compile(r"^NAS1149", re.IGNORECASE),   # Washers for inserts
    re.compile(r"^NAS16\d", re.IGNORECASE),   # NAS insert standards
    re.compile(r"^IOI", re.IGNORECASE),       # IOI inserts
    re.compile(r"^KN[LMC]", re.IGNORECASE),   # Keensert part numbers (KNL, KNM, KNC)
    re.compile(r"^402\d{6}", re.IGNORECASE),  # Rafael internal insert PNs (402xxxxxx)
    re.compile(r"^1084", re.IGNORECASE),      # Helicoil catalog numbers
]

# Thread pattern (e.g. "M4*1.5D", "4-40", "10-32", "1/4-28")
_THREAD_PATTERN = re.compile(
    r"M\d+\s*[*x×]\s*\d|"            # Metric: M4*1.5D, M6×2D
    r"\d+-\d+\s*(UNC|UNF|UNJC)?|"    # Imperial: 4-40, 10-32, 6-32UNC
    r"\d/\d+-\d+",                     # Fractional: 1/4-28
    re.IGNORECASE,
)


def _is_real_insert(item: Dict[str, Any]) -> bool:
    """
    Determine whether a BOM item is a genuine insert/hardware.

    Returns True if the item's description or cat_no matches known
    insert patterns. Returns False for sub-assembly parts.
    """
    cat_no = str(item.get("cat_no") or "").strip().upper()
    desc = str(item.get("description") or "").strip().upper()
    combined = f"{cat_no} {desc}"

    # Check description keywords
    for kw in _INSERT_DESC_KEYWORDS:
        if kw in combined:
            return True

    # Check part number patterns (cat_no or in description)
    for pattern in _INSERT_PN_PATTERNS:
        if pattern.search(cat_no) or pattern.search(desc):
            return True

    # Check thread pattern + context (thread description implies insert)
    if _THREAD_PATTERN.search(combined):
        # Thread dimensions alone are a strong signal for inserts
        # but only if the description doesn't clearly indicate a non-insert part
        non_insert_keywords = {
            "PLATE", "BRACKET", "HOUSING", "COVER", "FRAME",
            "PANEL", "ASSY", "ASSEMBLY", "BODY", "BASE",
            "SHIELD", "CAP", "RING", "SPACER", "ADAPTER",
        }
        if not any(nk in combined for nk in non_insert_keywords):
            return True

    return False


def validate_inserts_hardware(
    inserts_list: List[Dict[str, Any]],
    part_number: str = "",
) -> List[Dict[str, Any]]:
    """
    Filter inserts_hardware list, keeping only real inserts.

    Parameters
    ----------
    inserts_list : list[dict]
        Raw inserts_hardware from AI extraction (each has cat_no, qty, description).
    part_number : str
        The parent part number (for logging context).

    Returns
    -------
    list[dict]
        Only the items that are genuine inserts/hardware.
    """
    if not inserts_list:
        return inserts_list

    validated: list[dict] = []
    dropped: list[str] = []

    for item in inserts_list:
        if not isinstance(item, dict):
            continue
        if _is_real_insert(item):
            validated.append(item)
        else:
            cat_no = item.get("cat_no", "?")
            desc = item.get("description", "?")
            dropped.append(f"{cat_no} ({desc})")

    if dropped:
        logger.info(
            f"[INSERT-VALIDATOR] {part_number}: dropped {len(dropped)} non-insert BOM items: "
            f"{', '.join(dropped)}"
        )

    if len(validated) < len(inserts_list):
        logger.info(
            f"[INSERT-VALIDATOR] {part_number}: kept {len(validated)}/{len(inserts_list)} real inserts"
        )

    return validated
