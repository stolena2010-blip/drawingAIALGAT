"""
Config Manager — Read/write automation_config.json
"""
import json
from pathlib import Path
from typing import Dict, Any

CONFIG_PATH = Path(__file__).resolve().parent.parent.parent / "automation_config.json"
STATE_PATH = Path(__file__).resolve().parent.parent.parent / "automation_state.json"


def load_config() -> Dict[str, Any]:
    """Load automation config from JSON file."""
    if not CONFIG_PATH.exists():
        return _default_config()
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return _default_config()


def save_config(cfg: Dict[str, Any]) -> None:
    """Save automation config to JSON file."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_state() -> Dict[str, Any]:
    """Load automation state."""
    if not STATE_PATH.exists():
        return {}
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def reset_state() -> None:
    """Delete automation state file (reprocess all)."""
    if STATE_PATH.exists():
        STATE_PATH.unlink()


def _default_config() -> Dict[str, Any]:
    return {
        "shared_mailbox": "",
        "shared_mailboxes": [],
        "folder_name": "Inbox",
        "rerun_folder_name": "",
        "rerun_mailbox": "",
        "skip_senders": [],
        "skip_categories": [],
        "scan_from_date": "",
        "recipient_email": "",
        "download_root": "",
        "tosend_folder": "",
        "output_copy_folder": "",
        "poll_interval_minutes": 10,
        "max_messages": 200,
        "max_files_per_email": 15,
        "max_file_size_mb": 100,
        "stage1_skip_retry_resolution_px": 8000,
        "max_image_dimension": 4096,
        "recursive": True,
        "enable_retry": True,
        "auto_start": False,
        "auto_send": False,
        "archive_full": False,
        "cleanup_download": True,
        "mark_as_processed": True,
        "mark_category_name": "AI Processed",
        "mark_category_color": "preset20",
        "nodraw_category_name": "NO DRAW",
        "nodraw_category_color": "preset1",
        "heavy_category_name": "AI HEAVY",
        "heavy_category_color": "preset4",
        "confidence_level": "LOW",
        "debug_mode": False,
        "iai_top_red_fallback": True,
        "max_retries": 3,
        "scan_dpi": 200,
        "log_max_size_mb": 1,
        "usd_to_ils_rate": 3.7,
        "selected_stages": {str(i): True for i in range(1, 10)},
        "stage_models": {},
    }
