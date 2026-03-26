"""
Email Helpers — Shared mailbox connection and folder loading for Streamlit
This handles the shared mailbox ↔ folder loading logic that the automation
panel needs, using the existing GraphAPIHelper backend.
"""
import sys
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env")

from src.services.email.graph_helper import GraphAPIHelper


def test_mailbox_connection(mailbox: str) -> Tuple[bool, str]:
    """
    Test connection to a shared mailbox.
    Returns (success, message).
    """
    mailbox = (mailbox or "").strip()
    if not mailbox:
        return False, "לא הוגדרה תיבה"
    try:
        helper = GraphAPIHelper(shared_mailbox=mailbox)
        if helper.test_connection():
            return True, f"✅ חיבור תקין: {mailbox}"
        return False, f"❌ חיבור נכשל: {helper.last_error}"
    except Exception as e:
        return False, f"❌ שגיאה: {e}"


def test_all_mailboxes(mailboxes: List[str]) -> Dict[str, bool]:
    """Test connection to multiple mailboxes. Returns {mailbox: success}."""
    results = {}
    for mb in mailboxes:
        mb = mb.strip()
        if not mb:
            continue
        success, _ = test_mailbox_connection(mb)
        results[mb] = success
    return results


def load_folders_for_mailbox(mailbox: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    Load all folders (recursive) for a shared mailbox.
    Returns (folder_list, error_message).
    Each folder dict has: path, displayName, totalItemCount, id
    """
    mailbox = (mailbox or "").strip()
    if not mailbox:
        return [], "לא הוגדרה תיבה"

    try:
        helper = GraphAPIHelper(shared_mailbox=mailbox)
        if not helper.test_connection():
            return [], f"חיבור נכשל: {helper.last_error}"

        all_folders = helper.mailbox.list_folders_recursive() or []

        # Build deduplicated list with paths
        seen = set()
        result = []
        for folder in all_folders:
            path = (folder.get("path") or folder.get("displayName") or "").strip()
            if not path:
                continue
            key = path.lower()
            if key in seen:
                continue
            seen.add(key)
            result.append({
                "path": path,
                "displayName": folder.get("displayName", path),
                "totalItemCount": folder.get("totalItemCount"),
                "id": folder.get("id", ""),
            })

        # Ensure Inbox is present
        if "inbox" not in seen:
            wkn_inbox = helper.mailbox.get_well_known_folder("Inbox")
            if wkn_inbox:
                result.append({
                    "path": "Inbox",
                    "displayName": "Inbox",
                    "totalItemCount": wkn_inbox.get("totalItemCount"),
                    "id": wkn_inbox.get("id", ""),
                })

        result.sort(key=lambda x: x["path"].lower())
        return result, ""

    except Exception as e:
        return [], f"שגיאה בטעינת תיקיות: {e}"


def format_folder_label(path: str, total_item_count: Any = None) -> str:
    """Format folder path with item count for display."""
    clean_path = (path or "").strip()
    if not clean_path:
        return ""
    try:
        if total_item_count is not None and str(total_item_count).strip():
            count = int(total_item_count)
            if count >= 0:
                return f"{clean_path} ({count})"
    except (ValueError, TypeError):
        pass
    return clean_path


def parse_mailboxes_text(raw_text: str) -> List[str]:
    """Parse comma/semicolon/newline separated mailbox addresses."""
    if not raw_text:
        return []
    normalized = raw_text.replace(";", ",").replace("\n", ",")
    seen = set()
    result = []
    for part in normalized.split(","):
        mailbox = part.strip()
        if not mailbox:
            continue
        key = mailbox.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(mailbox)
    return result
