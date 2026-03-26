import json
import os
import re
import signal
import threading
import time
import shutil
import inspect
import zipfile
from uuid import uuid4
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, Any, Optional, List

from filelock import FileLock, Timeout as FileLockTimeout

from src.services.email.graph_helper import GraphAPIHelper
from customer_extractor_v3_dual import scan_folder
from src.utils.logger import get_logger

logger = get_logger(__name__)

# ─── Stdout redirector (mirrors Tkinter LogRedirector) ────────────────
_STATUS_LOG_PATH = Path.cwd() / "status_log.txt"
_ANSI_RE = re.compile(r'\x1b\[[0-9;]*m')


class _StatusLogRedirector:
    """Redirects stdout to both the original console AND status_log.txt.
    This gives the Streamlit log panel the same detailed output that
    Tkinter showed via its LogRedirector (print-based stage costs,
    dimensions, rotation checks, etc.)."""

    def __init__(self, original_stdout):
        self.original = original_stdout
        self._lock = threading.Lock()

    def write(self, message) -> None:
        try:
            self.original.write(message)
        except Exception:
            pass
        if message and message.strip():
            try:
                clean = _ANSI_RE.sub('', message.rstrip())
                with self._lock:
                    with open(_STATUS_LOG_PATH, "a", encoding="utf-8") as f:
                        # Logger lines already have HH:MM:SS timestamp — write as-is
                        if (
                            len(clean) > 8
                            and clean[2] == ':' and clean[5] == ':'
                            and clean[8] in (' ', '\u2502')
                        ):
                            f.write(f"{clean}\n")
                        else:
                            ts = datetime.now().strftime("%H:%M:%S")
                            f.write(f"[{ts}] {clean}\n")
            except Exception:
                pass

    def flush(self) -> None:
        try:
            self.original.flush()
        except Exception:
            pass


def _trim_status_log() -> None:
    """Keep status_log.txt under 5000 lines."""
    try:
        if _STATUS_LOG_PATH.exists():
            lines = _STATUS_LOG_PATH.read_text(encoding="utf-8").splitlines()
            if len(lines) > 5500:
                _STATUS_LOG_PATH.write_text(
                    "\n".join(lines[-5000:]) + "\n", encoding="utf-8"
                )
    except Exception:
        pass

# Drawing file extensions (matching customer_extractor_v3_dual.DRAWING_EXTS + extras)
_DRAWING_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}


def _find_unrar() -> str | None:
    """Locate UnRAR.exe on the system."""
    for p in (
        Path(r"C:\Program Files\WinRAR\UnRAR.exe"),
        Path(r"C:\Program Files (x86)\WinRAR\UnRAR.exe"),
    ):
        if p.exists():
            return str(p)
    # Check PATH
    import shutil
    return shutil.which("UnRAR")


def _count_drawing_files(message_dir: Path) -> int:
    """Count drawing files in a message directory, including files inside ZIPs and RARs."""
    count = 0
    for f in message_dir.rglob("*"):
        if not f.is_file():
            continue
        ext = f.suffix.lower()
        if ext in _DRAWING_EXTS:
            count += 1
        elif ext == ".zip":
            try:
                if zipfile.is_zipfile(f):
                    with zipfile.ZipFile(f, "r") as zf:
                        for member in zf.namelist():
                            if member.endswith("/"):
                                continue
                            member_ext = Path(member).suffix.lower()
                            if member_ext in _DRAWING_EXTS:
                                count += 1
            except Exception:
                pass  # corrupt zip — will be handled later
        elif ext == ".rar":
            # Peek inside RAR using UnRAR l (list) — no extraction needed
            try:
                unrar = _find_unrar()
                if unrar:
                    import subprocess
                    result = subprocess.run(
                        [unrar, 'lb', str(f)],
                        capture_output=True, text=True, timeout=15,
                        encoding='utf-8', errors='replace',
                    )
                    if result.returncode == 0:
                        for line in result.stdout.splitlines():
                            member_ext = Path(line.strip()).suffix.lower()
                            if member_ext in _DRAWING_EXTS:
                                count += 1
                else:
                    # No UnRAR — optimistically assume RAR contains drawings
                    count += 1
            except Exception:
                count += 1  # assume it has drawings to avoid false NO DRAW
    return count


def _now_iso() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"



# File renaming is now handled in customer_extractor_v3_dual.py via _copy_folder_to_tosend()


# ── Outlook category preset → CSS color mapping ──────────────────────
_PRESET_COLORS = {
    "preset0": ("#FFFFFF", "#333333"),   # None (no color)
    "preset1": ("#E74C3C", "#FFFFFF"),   # Red
    "preset2": ("#F39C12", "#FFFFFF"),   # Orange
    "preset3": ("#8B6914", "#FFFFFF"),   # Brown
    "preset4": ("#F1C40F", "#333333"),   # Yellow
    "preset5": ("#27AE60", "#FFFFFF"),   # Green
    "preset6": ("#1ABC9C", "#FFFFFF"),   # Teal
    "preset7": ("#808000", "#FFFFFF"),   # Olive
    "preset8": ("#2980B9", "#FFFFFF"),   # Blue
    "preset9": ("#8E44AD", "#FFFFFF"),   # Purple
    "preset10": ("#E91E8C", "#FFFFFF"),  # Pink
    "preset11": ("#95A5A6", "#FFFFFF"),  # Gray
    "preset12": ("#922B21", "#FFFFFF"),  # Dark Red
    "preset13": ("#D35400", "#FFFFFF"),  # Dark Orange
    "preset14": ("#6B4226", "#FFFFFF"),  # Dark Brown
    "preset15": ("#B7950B", "#333333"),  # Dark Yellow
    "preset16": ("#1E8449", "#FFFFFF"),  # Dark Green
    "preset17": ("#148F77", "#FFFFFF"),  # Dark Teal
    "preset18": ("#556B2F", "#FFFFFF"),  # Dark Olive
    "preset19": ("#1A5276", "#FFFFFF"),  # Dark Blue
    "preset20": ("#6C3483", "#FFFFFF"),  # Dark Purple
    "preset21": ("#C2185B", "#FFFFFF"),  # Dark Pink
    "preset22": ("#616A6B", "#FFFFFF"),  # Dark Gray
    "preset23": ("#1C2833", "#FFFFFF"),  # Black
    "preset24": ("#D5D8DC", "#333333"),  # Light Gray
    "preset25": ("#85C1E9", "#333333"),  # Light Blue
}


def _build_category_banner(categories: list, category_color_map: dict = None) -> str:
    """
    Build an HTML banner showing the original message's Outlook categories.
    Each category becomes a colored pill/tag in the banner.
    Returns empty string if no categories.
    """
    if not categories:
        return ""

    if category_color_map is None:
        category_color_map = {}

    pills = []
    for cat_name in categories:
        # Look up the preset color from the mailbox's master category list
        preset = category_color_map.get(cat_name, "")
        bg, fg = _PRESET_COLORS.get(preset, ("#7B68EE", "#FFFFFF"))
        pills.append(
            f'<span style="display:inline-block; background-color:{bg}; color:{fg}; '
            f'padding:3px 10px; border-radius:12px; margin:0 4px; font-size:13px; '
            f'font-weight:bold;">\U0001f3f7\ufe0f {cat_name}</span>'
        )

    return (
        '<div dir="rtl" style="padding:8px 10px; margin-bottom:8px; '
        'background-color:#F0EBF8; border-right:4px solid #9B59B6; font-size:13px;">'
        '\U0001f3f7\ufe0f <b>קטגוריות מקוריות:</b> '
        + " ".join(pills)
        + "</div>"
    )


def _load_json(path: Path, default: Dict[str, Any]) -> Dict[str, Any]:
    if not path.exists():
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.debug(f"Fallback: {e}")
        return default


class _SetEncoder(json.JSONEncoder):
    """JSON encoder that converts sets to sorted lists."""
    def default(self, obj):
        if isinstance(obj, set):
            return sorted(obj)
        return super().default(obj)


def _save_json(path: Path, data: Dict[str, Any]) -> None:
    """Save JSON atomically (write temp then replace)."""
    tmp = path.with_suffix(".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, cls=_SetEncoder)
        tmp.replace(path)
        
        # Verify save succeeded
        ids_count = len(data.get("processed_ids", []))
        mailboxes = data.get("processed_ids_by_mailbox", {})
        mb_counts = {k: len(v) for k, v in mailboxes.items()}
        logger.info(f"\U0001f4be State saved: {ids_count} total IDs, mailboxes: {mb_counts}")
        
    except PermissionError as e:
        logger.error(f"\u274c PERMISSION ERROR saving state: {e}")
        print(f"\u274c PERMISSION ERROR: Cannot write {path} \u2014 is it open in another program?")
        # Fallback: try direct write (not atomic)
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2, cls=_SetEncoder)
            logger.info(f"\U0001f4be State saved (fallback direct write)")
        except Exception as e2:
            logger.error(f"\u274c FALLBACK ALSO FAILED: {e2}")
            print(f"\u274c CRITICAL: Cannot save state at all: {e2}")
    except Exception as e:
        logger.error(f"\u274c ERROR saving state to {path}: {e}")
        print(f"WARNING: Failed to save state: {e}")
        # Fallback: try direct write
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2, cls=_SetEncoder)
            logger.info(f"\U0001f4be State saved (fallback direct write)")
        except Exception as e2:
            logger.error(f"\u274c FALLBACK ALSO FAILED: {e2}")
        try:
            if tmp.exists():
                tmp.unlink()
        except Exception:
            pass


def _rotate_log_if_needed(log_path: Path, max_size_bytes: int = 1_000_000) -> None:
    """Rotate log file if it exceeds max_size_bytes (default 1MB)."""
    try:
        if log_path.exists() and log_path.stat().st_size > max_size_bytes:
            timestamp = datetime.now().strftime("%Y%m%d")
            rotated = log_path.parent / f"automation_log_{timestamp}.jsonl"
            # Avoid overwriting existing rotated file
            counter = 1
            while rotated.exists():
                rotated = log_path.parent / f"automation_log_{timestamp}_{counter}.jsonl"
                counter += 1
            log_path.rename(rotated)
            logger.info(f"[LOG] Rotated {log_path.name} -> {rotated.name} ({log_path.stat if False else 'size exceeded'})")
    except Exception as e:
        logger.error(f"[LOG] Rotation failed: {e}")


def _append_log(path: Path, data: Dict[str, Any]) -> None:
    try:
        _rotate_log_if_needed(path)
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(data, ensure_ascii=False) + "\n")
    except Exception as e:
        logger.debug(f"Ignored: {e}")
        pass


def _clean_sender_line(line: str) -> str:
    return line.replace("כתובת שולח:", "").replace("From:", "").strip()


def _scan_folder_compat(
    message_dir: Path,
    recursive: bool,
    stages: Dict[int, bool],
    enable_retry: bool,
    tosend_folder: Optional[str],
    confidence_level: str,
    stage1_skip_retry_resolution_px: int,
    max_file_size_mb: int,
    max_image_dimension: int,
):
    kwargs = {
        "recursive": recursive,
        "selected_stages": stages,
        "enable_image_retry": enable_retry,
        "tosend_folder": tosend_folder,
        "confidence_level": confidence_level,
    }

    try:
        sig = inspect.signature(scan_folder)
        if "stage1_skip_retry_resolution_px" in sig.parameters:
            kwargs["stage1_skip_retry_resolution_px"] = stage1_skip_retry_resolution_px
        if "max_file_size_mb" in sig.parameters:
            kwargs["max_file_size_mb"] = max_file_size_mb
        if "max_image_dimension" in sig.parameters:
            kwargs["max_image_dimension"] = max_image_dimension
    except Exception as e:
        logger.debug(f"Ignored: {e}")
        pass

    return scan_folder(message_dir, **kwargs)


class AutomationRunner:
    def __init__(
        self,
        config_path: Path,
        state_path: Path,
        status_callback: Optional[Callable[[str], None]] = None
    ) -> None:
        self.config_path = config_path
        self.state_path = state_path
        self.status_callback = status_callback
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None

    def _status(self, message: str) -> None:
        # Write to status_log.txt so Streamlit (and any other UI) can read it
        try:
            from datetime import datetime as _dt
            ts = _dt.now().strftime("%H:%M:%S")
            log_path = Path.cwd() / "status_log.txt"
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{ts}] {message}\n")
            # Trim to last 5000 lines periodically
            try:
                lines = log_path.read_text(encoding="utf-8").splitlines()
                if len(lines) > 5500:
                    log_path.write_text("\n".join(lines[-5000:]) + "\n", encoding="utf-8")
            except Exception:
                pass
        except Exception:
            pass
        if self.status_callback:
            try:
                self.status_callback(message)
            except Exception as e:
                logger.debug(f"Ignored: {e}")
                pass

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def start_with_signals(self) -> None:
        """Start with OS signal handling (for server/service mode)."""
        def _handle_signal(signum, frame) -> None:
            sig_name = signal.Signals(signum).name
            self._status(f"Received {sig_name} \u2014 finishing current task then stopping...")
            self._stop_event.set()
        
        signal.signal(signal.SIGTERM, _handle_signal)
        signal.signal(signal.SIGINT, _handle_signal)
        
        self.start()
        
        # Block main thread until stopped
        try:
            while self._thread and self._thread.is_alive():
                self._thread.join(timeout=1)
        except KeyboardInterrupt as e:
            logger.debug(f"Error: {e}")
            self._status("KeyboardInterrupt \u2014 stopping...")
            self._stop_event.set()

    def stop(self) -> None:
        self._stop_event.set()
        self._write_health("stopped")

    def _handle_rerun(self, helper, msg_id: str, mailbox: str, 
                      recipient: str, download_root: Path, config: dict) -> bool:
        """
        Handle RERUN: download our own Reply, swap ALL_B2B→B2B-0_, resend.
        Returns True if ALL_B2B was found and sent. False if not found (e.g., confidence=LOW).
        """
        run_stamp = datetime.now().strftime("%Y%m%d%H%M%S")
        run_dir = download_root / f"rerun_{run_stamp}_{msg_id[:8]}"
        
        # Step 1: Download the dragged-back Reply
        download_result = helper.download_message_by_id(msg_id, str(run_dir))
        if not download_result.get("success"):
            logger.error(f"RERUN: Failed to download message")
            return False
        
        message_dir = Path(download_result.get("message_dir"))
        original_body_html = download_result.get("original_body_html", "")
        original_subject = download_result.get("subject", "RERUN")
        sender_from_body = download_result.get("sender", "")
        original_categories = download_result.get("categories", [])

        # Fetch category→color map from mailbox (for colored banner)
        category_color_map = {}
        if original_categories:
            try:
                category_color_map = helper.mailbox.list_categories()
            except Exception:
                pass
        
        # Step 2: Find ALL_B2B file in attachments
        b2b_all_files = list(message_dir.rglob("ALL_B2B*.txt"))
        if not b2b_all_files:
            logger.info(f"RERUN: No ALL_B2B file — probably confidence=LOW (all rows already sent)")
            return False
        
        b2b_all_file = b2b_all_files[0]
        logger.info(f"RERUN: Found {b2b_all_file.name}")
        
        # Step 3: Remove old B2B-0_ file(s)
        for old in message_dir.rglob("B2B-0_*.txt"):
            try:
                old.unlink()
                logger.info(f"RERUN: Removed {old.name}")
            except Exception:
                pass
        # Also remove B2B- files that aren't ALL_B2B
        for old in message_dir.rglob("B2B-*.txt"):
            if not old.name.startswith("ALL_B2B"):
                try:
                    old.unlink()
                except Exception:
                    pass
        
        # Step 4: Rename ALL_B2B → B2B-0_
        # ALL_B2B-0_200002-PO123.txt → B2B-0_200002-PO123.txt
        new_name = b2b_all_file.name.replace("ALL_B2B-", "B2B-").replace("ALL_B2B_", "B2B-")
        new_path = b2b_all_file.parent / new_name
        try:
            b2b_all_file.rename(new_path)
            logger.info(f"RERUN: {b2b_all_file.name} → {new_name}")
        except Exception as e:
            logger.error(f"RERUN: Rename failed: {e}")
            return False
        
        # Step 5: Swap ALL_METADATA → metadata
        for meta_all in list(message_dir.rglob("ALL_METADATA.json")) + list(message_dir.rglob("metadata_ALL.json")):
            meta_target = meta_all.parent / "metadata.json"
            try:
                if meta_target.exists():
                    meta_target.unlink()
                meta_all.rename(meta_target)
                logger.info(f"RERUN: {meta_all.name} → metadata.json")
            except Exception as e:
                logger.warning(f"RERUN: metadata swap failed: {e}")
        
        # Step 6: Collect files to send (skip internal files)
        skip_prefixes = ('file_classification_', 'drawing_results_', 'SUMMARY_',
                         'email.txt', 'ALL_METADATA', 'metadata_ALL', 'ALL_B2B', 'B2B_ALL')
        
        files_to_send = []
        search_dir = message_dir
        # Check if there's a subfolder with the actual files
        subdirs = [d for d in message_dir.iterdir() if d.is_dir()]
        if subdirs:
            search_dir = subdirs[0]
        
        for f in search_dir.iterdir():
            if f.is_file() and f.suffix.lower() != '.zip':
                if not any(f.name.startswith(sp) for sp in skip_prefixes):
                    files_to_send.append(f)
        
        if not files_to_send:
            logger.warning(f"RERUN: No files to send")
            return False
        
        # Step 7: Build subject (keep original) and body
        email_subject = original_subject
        rerun_notice = '<div dir="rtl" style="background-color:#E8D5F5; padding:10px; margin-bottom:10px; border-right:4px solid #7B2D8E; font-size:14px;">🔄 <b>מייל חוזר לאחר אישור משתמש לכלול את כל הפריטים</b></div>'

        # Build category banner from original message categories
        category_banner = _build_category_banner(original_categories, category_color_map)

        email_body = rerun_notice + category_banner + (original_body_html or "")
        
        # Step 8: Send
        logger.info(f"RERUN: Sending {len(files_to_send)} files to {recipient}")
        try:
            helper.send_email(
                to_address=recipient,
                subject=email_subject,
                body=email_body,
                attachments=files_to_send,
                body_type="HTML",
                replace_display_with_filename=True,
            )
        except Exception as e:
            logger.error(f"RERUN: Send failed: {e}")
            return False
        
        # Step 9: Log
        _append_log(
            self.state_path.parent / "automation_log.jsonl",
            {
                "id": str(uuid4()),
                "shared_mailbox": mailbox,
                "message_id": msg_id,
                "sender": sender_from_body,
                "type": "RERUN",
                "files_sent": len(files_to_send),
                "cost_usd": 0.0,
                "timestamp": _now_iso(),
            }
        )
        
        return True

    def _check_and_handle_pending_reruns(
        self, helper, mailbox: str, folder_id: str, recipient: str, 
        download_root: Path, config: dict,
        mailbox_processed_ids: set, mark_as_processed: bool,
        cleanup_download: bool
    ) -> int:
        """
        Quick check for pending RERUN emails and handle them immediately.
        Called after each normal email processing.
        Returns number of RERUNs handled.
        """
        rerun_count = 0
        try:
            # Fetch latest messages (light API call — no AI cost)
            messages = helper.mailbox.list_messages(folder_id, limit=20)
            if not messages:
                return 0
            
            for msg in messages:
                msg_id = msg.get("id", "")
                if not msg_id or msg_id in mailbox_processed_ids:
                    continue
                
                # Check if sender is ourselves (= RERUN)
                msg_sender = (
                    msg.get("from", {}).get("emailAddress", {}).get("address", "") or ""
                ).lower().strip()
                
                if msg_sender != mailbox.lower().strip():
                    continue  # Not a RERUN
                
                logger.info(f"🔄 RERUN (priority): detected between emails — {msg_id[:8]}")
                self._status(f"🔄 RERUN מיידי — החלפת B2B...")
                
                try:
                    rerun_ok = self._handle_rerun(
                        helper=helper,
                        msg_id=msg_id,
                        mailbox=mailbox,
                        recipient=recipient,
                        download_root=download_root,
                        config=config,
                    )
                    if rerun_ok:
                        logger.info(f"✅ RERUN (priority) completed — {msg_id[:8]}")
                        rerun_count += 1
                    else:
                        logger.warning(f"⚠️ RERUN (priority) skipped — no ALL_B2B")
                except Exception as e:
                    logger.error(f"❌ RERUN (priority) error: {e}")
                
                # Mark as processed
                if mark_as_processed:
                    try:
                        helper.ensure_category("AI RERUN", "preset9")
                        helper.mark_message_processed(msg_id, "AI RERUN")
                    except Exception:
                        pass
                
                mailbox_processed_ids.add(msg_id)
                
                if cleanup_download:
                    try:
                        for d in download_root.glob(f"rerun_*_{msg_id[:8]}"):
                            if d.is_dir():
                                shutil.rmtree(d, ignore_errors=True)
                    except Exception:
                        pass
        
        except Exception as e:
            logger.debug(f"RERUN priority check failed (non-critical): {e}")
        
        return rerun_count

    def _scan_rerun_folder(
        self, helper, mailbox: str, all_folders: list,
        rerun_folder_name: str, recipient: str,
        download_root: Path, config: dict,
        mailbox_processed_ids: set, mark_as_processed: bool,
    ) -> int:
        """
        Scan the dedicated RERUN folder for unprocessed messages.
        Lightweight: rename ALL_B2B → B2B-0_ + send, no AI cost.
        Returns number of RERUNs handled.
        """
        if not rerun_folder_name:
            return 0

        rerun_folder_id = None
        rerun_folder_lower = rerun_folder_name.strip().lower()

        for folder in all_folders:
            folder_path = str(folder.get("path", "")).strip().lower()
            folder_display = str(folder.get("displayName", "")).strip().lower()
            if folder_path == rerun_folder_lower or folder_display == rerun_folder_lower:
                rerun_folder_id = folder.get("id")
                break

        if not rerun_folder_id:
            # Try flat search
            for folder in (helper.mailbox.list_folders(show_inbox_only=False) or []):
                if folder.get("displayName", "").lower() == rerun_folder_lower:
                    rerun_folder_id = folder.get("id")
                    break

        if not rerun_folder_id:
            # Fallback: search inside well-known Inbox (handles Hebrew "תיבת דואר נכנס")
            inbox_info = helper.mailbox.get_well_known_folder("Inbox")
            if inbox_info and inbox_info.get("id"):
                children = helper.mailbox._list_child_folders_all(inbox_info["id"])
                for child in children:
                    if str(child.get("displayName", "")).strip().lower() == rerun_folder_lower:
                        rerun_folder_id = child.get("id")
                        logger.info(f"📂 Found RERUN folder inside Inbox (well-known): {child.get('displayName')}")
                        break

        if not rerun_folder_id:
            logger.info(f"RERUN folder '{rerun_folder_name}' not found — skipping")
            return 0

        rerun_count = 0
        try:
            rerun_messages = helper.mailbox.list_messages(rerun_folder_id, limit=50, received_after=None)
            rerun_new = [m for m in rerun_messages if m.get("id") and m.get("id") not in mailbox_processed_ids]

            if not rerun_new:
                return 0

            logger.info(f"🔄 Found {len(rerun_new)} messages in RERUN folder '{rerun_folder_name}'")

            for ridx, rmsg in enumerate(rerun_new, 1):
                rmsg_id = rmsg.get("id")
                if not rmsg_id:
                    continue

                logger.info(f"🔄 RERUN (from folder): {ridx}/{len(rerun_new)}")
                self._status(f"🔄 RERUN {ridx}/{len(rerun_new)} — החלפת B2B...")

                try:
                    rerun_ok = self._handle_rerun(
                        helper=helper,
                        msg_id=rmsg_id,
                        mailbox=mailbox,
                        recipient=recipient,
                        download_root=download_root,
                        config=config,
                    )
                    if rerun_ok:
                        logger.info(f"✅ RERUN (folder) completed — ALL_B2B sent")
                        rerun_count += 1
                    else:
                        logger.warning(f"⚠️ RERUN (folder) skipped — no ALL_B2B found")
                except Exception as e:
                    logger.error(f"❌ RERUN (folder) error: {e}")

                # Mark processed
                if mark_as_processed:
                    try:
                        helper.ensure_category("AI RERUN", "preset9")
                        helper.mark_message_processed(rmsg_id, "AI RERUN")
                    except Exception:
                        pass

                mailbox_processed_ids.add(rmsg_id)
        except Exception as e:
            logger.debug(f"RERUN folder scan error (non-critical): {e}")

        return rerun_count

    def run_once(self) -> None:
        self._stop_event.clear()
        self._activate_stdout_redirect()
        try:
            self._run_once_internal()
        finally:
            self._deactivate_stdout_redirect()

    def run_heavy(self) -> None:
        """Process only emails marked 'AI HEAVY' — no file-count threshold."""
        self._stop_event.clear()
        self._activate_stdout_redirect()
        try:
            self._run_once_internal(heavy_only=True)
        finally:
            self._deactivate_stdout_redirect()

    def _activate_stdout_redirect(self) -> None:
        """Redirect stdout → status_log.txt and point logging StreamHandlers
        through the redirector so ALL output (print + logger) lands in the file."""
        import sys, logging
        from logging.handlers import RotatingFileHandler as _RFH
        if not isinstance(sys.stdout, _StatusLogRedirector):
            self._original_stdout = sys.stdout
            sys.stdout = _StatusLogRedirector(sys.stdout)
        # Re-point console StreamHandlers through the redirector
        for h in logging.getLogger().handlers[:]:
            if isinstance(h, logging.StreamHandler) and not isinstance(h, (logging.FileHandler, _RFH)):
                h.stream = sys.stdout

    def _deactivate_stdout_redirect(self) -> None:
        """Restore original stdout and StreamHandler streams."""
        import sys, logging
        from logging.handlers import RotatingFileHandler as _RFH
        if hasattr(self, '_original_stdout') and self._original_stdout:
            for h in logging.getLogger().handlers[:]:
                if isinstance(h, logging.StreamHandler) and not isinstance(h, (logging.FileHandler, _RFH)):
                    h.stream = self._original_stdout
            sys.stdout = self._original_stdout
            self._original_stdout = None

    def _loop(self) -> None:
        self._activate_stdout_redirect()
        try:
            self._in_loop = True
            self._write_health("running", "loop started")
            while not self._stop_event.is_set():
                self._write_health("processing")
                self._run_once_internal()
                _trim_status_log()
                self._write_health("waiting")
                config = _load_json(self.config_path, {})
                interval_minutes = max(int(config.get("poll_interval_minutes", 10)), 1)
                for _ in range(interval_minutes * 60):
                    if self._stop_event.is_set():
                        break
                    time.sleep(1)
        finally:
            self._in_loop = False
            self._deactivate_stdout_redirect()

    def _write_health(self, status: str, details: str = "") -> None:
        """Write health status file for external monitoring."""
        health_path = self.state_path.parent / "health.json"
        try:
            health = {
                "status": status,
                "timestamp": datetime.now().isoformat(),
                "details": details,
                "pid": os.getpid(),
            }
            health_path.write_text(json.dumps(health, indent=2), encoding="utf-8")
        except Exception as e:
            logger.debug(f"Ignored: {e}")
            pass

    def _export_cycle_report_if_configured(self, config: dict, heavy_only: bool, cycle_start: datetime | None = None) -> None:
        """Export/update scheduler Excel report when output folder is configured."""
        report_folder = str(config.get("scheduler_report_folder", "") or "").strip()
        if not report_folder:
            return
        run_type = "heavy" if heavy_only else "regular"
        try:
            from streamlit_app.backend.report_exporter import export_schedule_report
            today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            saved_path = export_schedule_report(report_folder, run_type, today_start)
            logger.info(f"📊 Scheduler report saved: {saved_path}")
            self._status(f"📊 דוח נשמר: {saved_path}")
        except Exception as exc:
            logger.error(f"📊 Failed to save scheduler report: {exc}")
            self._status(f"📊 שגיאה בשמירת דוח Excel: {exc}")

    def _notify_failure(self, error_msg: str) -> None:
        """Write failure notification to file (for external monitor pickup)."""
        alert_path = self.state_path.parent / "alert.json"
        try:
            alert = {
                "type": "failure",
                "message": error_msg,
                "timestamp": datetime.now().isoformat(),
                "pid": os.getpid(),
            }
            alert_path.write_text(json.dumps(alert, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            logger.debug(f"Ignored: {e}")
            pass

    @staticmethod
    def validate_config() -> List[str]:
        """Validate .env and return list of errors (empty = OK)."""
        from dotenv import dotenv_values

        errors: List[str] = []
        env = dotenv_values()

        if not env.get("AZURE_OPENAI_ENDPOINT"):
            errors.append("AZURE_OPENAI_ENDPOINT חסר")
        if not env.get("AZURE_OPENAI_API_KEY"):
            errors.append("AZURE_OPENAI_API_KEY חסר")
        if not env.get("AZURE_OPENAI_DEPLOYMENT"):
            errors.append("AZURE_OPENAI_DEPLOYMENT חסר")

        if not env.get("GRAPH_TENANT_ID"):
            errors.append("GRAPH_TENANT_ID חסר")
        if not env.get("GRAPH_CLIENT_ID"):
            errors.append("GRAPH_CLIENT_ID חסר")
        if not env.get("GRAPH_CLIENT_SECRET"):
            errors.append("GRAPH_CLIENT_SECRET חסר")

        endpoint = env.get("AZURE_OPENAI_ENDPOINT", "")
        if endpoint and not str(endpoint).startswith("https://"):
            errors.append(f"AZURE_OPENAI_ENDPOINT לא תקין: {endpoint}")

        return errors

    @staticmethod
    def _get_configured_mailboxes(config: Dict[str, Any]) -> List[str]:
        seen = set()
        result: List[str] = []

        configured = config.get("shared_mailboxes")
        if isinstance(configured, list):
            candidates = [str(x).strip() for x in configured]
        elif isinstance(configured, str):
            normalized = configured.replace(";", ",").replace("\n", ",")
            candidates = [part.strip() for part in normalized.split(",")]
        else:
            candidates = []

        fallback_single = str(config.get("shared_mailbox", "")).strip()
        if fallback_single:
            candidates.append(fallback_single)

        for mailbox in candidates:
            if not mailbox:
                continue
            key = mailbox.lower()
            if key in seen:
                continue
            seen.add(key)
            result.append(mailbox)

        return result

    def _run_once_internal(self, heavy_only: bool = False) -> None:
        # ═══ Cross-process file lock: prevent parallel runners ═══
        lock_path = self.state_path.with_suffix(".lock")
        lock = FileLock(lock_path, timeout=0)
        try:
            lock.acquire()
        except FileLockTimeout:
            logger.warning("⚠️ Another automation runner is already active — skipping this run")
            self._status("⚠️ ריצה כבר פעילה בתהליך אחר — מדלג")
            return
        try:
            self._run_once_internal_locked(heavy_only)
        finally:
            lock.release()

    def _run_once_internal_locked(self, heavy_only: bool = False) -> None:
        cycle_start = datetime.now()
        config = _load_json(self.config_path, {})
        if not hasattr(self, "_config_validated"):
            errors = self.validate_config()
            if errors:
                self._status(f"שגיאות הגדרה: {' | '.join(errors)}")
                return
            self._config_validated = True

        auto_send = bool(config.get("auto_send", False))
        archive_full = bool(config.get("archive_full", False))
        cleanup_download = bool(config.get("cleanup_download", True))
        mark_as_processed = bool(config.get("mark_as_processed", True))  # סימון מיילים מעובדים
        mark_category_name = config.get("mark_category_name", "AI Processed")
        mark_category_color = config.get("mark_category_color", "preset0")
        nodraw_category_name = config.get("nodraw_category_name", "NO DRAW")
        nodraw_category_color = config.get("nodraw_category_color", "preset1")
        heavy_category_name = config.get("heavy_category_name", "AI HEAVY")
        heavy_category_color = config.get("heavy_category_color", "preset4")
        try:
            max_files_per_email = max(int(config.get("max_files_per_email", 0) or 0), 0)
        except Exception:
            max_files_per_email = 0
        confidence_level = config.get("confidence_level", "LOW")  # B2B confidence level
        try:
            stage1_skip_retry_resolution_px = max(int(config.get("stage1_skip_retry_resolution_px", 8000) or 8000), 0)
        except Exception as e:
            logger.debug(f"Handled: {e}")
            stage1_skip_retry_resolution_px = 8000
        try:
            max_file_size_mb = max(int(config.get("max_file_size_mb", 100) or 100), 1)
        except Exception as e:
            logger.debug(f"Handled: {e}")
            max_file_size_mb = 100
        try:
            max_image_dimension = max(int(config.get("max_image_dimension", 3072) or 3072), 256)
        except Exception as e:
            logger.debug(f"Handled: {e}")
            max_image_dimension = 3072
        shared_mailboxes = self._get_configured_mailboxes(config)
        required = ["folder_name", "download_root"]
        if auto_send:
            required.append("recipient_email")
        if not shared_mailboxes or any(not config.get(k) for k in required):
            self._status("חסרות הגדרות חובה (תיבות, תיקייה, נמען, תיקיית הורדה)")
            return

        state = _load_json(
            self.state_path,
            {
                "processed_ids": [],
                "last_checked": None,
                "processed_ids_by_mailbox": {},
                "last_checked_by_mailbox": {}
            }
        )
        processed_ids_by_mailbox = state.get("processed_ids_by_mailbox", {})
        if not isinstance(processed_ids_by_mailbox, dict):
            processed_ids_by_mailbox = {}
        last_checked_by_mailbox = state.get("last_checked_by_mailbox", {})
        if not isinstance(last_checked_by_mailbox, dict):
            last_checked_by_mailbox = {}

        legacy_processed_ids = set(state.get("processed_ids", []))
        legacy_last_checked = state.get("last_checked")
        log_path = Path.cwd() / "automation_log.jsonl"

        # Apply advanced settings to environment
        os.environ["AI_DRAW_DEBUG"] = "true" if config.get("debug_mode", False) else "false"
        os.environ["IAI_TOP_RED_FALLBACK_ENABLED"] = "true" if config.get("iai_top_red_fallback", True) else "false"
        os.environ["USD_TO_ILS_RATE"] = str(config.get("usd_to_ils_rate", 3.7))

        # Apply per-stage model overrides from GUI config
        stage_models = config.get("stage_models") or {}
        if stage_models:
            for stage_n_str, model_name in stage_models.items():
                model_name = (model_name or "").strip()
                if model_name:
                    env_key = f"STAGE_{stage_n_str}_MODEL"
                    os.environ[env_key] = model_name
                    logger.info(f"🔧 Stage {stage_n_str} model override: {model_name}")

        # Rotate log if needed (using configured size)
        log_max_bytes = int(config.get("log_max_size_mb", 1)) * 1_000_000
        _rotate_log_if_needed(log_path, max_size_bytes=log_max_bytes)

        folder_name = config.get("folder_name", "Inbox")
        rerun_folder_name = config.get("rerun_folder_name") or None  # e.g., "AI RERUN"
        rerun_mailbox = config.get("rerun_mailbox", "").strip() or None  # separate mailbox for RERUN
        skip_senders = set(s.strip().lower() for s in (config.get("skip_senders") or []) if s.strip())
        if skip_senders:
            logger.info(f"🚫 Skip senders list ({len(skip_senders)}): {', '.join(sorted(skip_senders))}")
        skip_categories = set(c.strip() for c in (config.get("skip_categories") or []) if c.strip())
        if skip_categories:
            logger.info(f"🏷️ Skip categories list ({len(skip_categories)}): {', '.join(sorted(skip_categories))}")
        scan_from_date_str = config.get("scan_from_date", "").strip()
        scan_from_datetime = None
        if scan_from_date_str:
            try:
                # Support DD/MM/YYYY HH:MM format
                scan_from_datetime = datetime.strptime(scan_from_date_str, "%d/%m/%Y %H:%M")
                logger.info(f"📅 Scan from date: {scan_from_datetime.isoformat()}")
            except ValueError:
                try:
                    # Also support YYYY-MM-DD HH:MM
                    scan_from_datetime = datetime.strptime(scan_from_date_str, "%Y-%m-%d %H:%M")
                    logger.info(f"📅 Scan from date: {scan_from_datetime.isoformat()}")
                except ValueError:
                    logger.warning(f"⚠️ Invalid scan_from_date format: '{scan_from_date_str}' — ignoring")
        download_root = Path(config.get("download_root"))
        tosend_folder = config.get("tosend_folder") or None
        output_copy_folder = config.get("output_copy_folder") or None
        recipient = config.get("recipient_email")

        selected_stages = config.get("selected_stages") or {"1": True, "2": True, "3": True, "4": True, "5": True, "6": True, "7": True, "8": True}
        stages = {
            1: bool(selected_stages.get("1", True)),
            2: bool(selected_stages.get("2", True)),
            3: bool(selected_stages.get("3", True)),
            4: bool(selected_stages.get("4", True)),
            5: bool(selected_stages.get("5", True)),
            6: bool(selected_stages.get("6", True)),
            7: bool(selected_stages.get("7", True)),
            8: bool(selected_stages.get("8", True)),
        }

        recursive = bool(config.get("recursive", True))
        enable_retry = bool(config.get("enable_retry", True))
        limit = int(config.get("max_messages", 200))

        had_errors = False
        total_new_messages = 0

        try:
            # ═══ RERUN SCAN ON SEPARATE MAILBOX — runs FIRST ═══
            if rerun_folder_name and rerun_mailbox and rerun_mailbox.lower() not in [m.lower() for m in shared_mailboxes]:
                logger.info(f"🔄 Scanning RERUN folder in separate mailbox: {rerun_mailbox}")
                try:
                    rerun_helper = GraphAPIHelper(shared_mailbox=rerun_mailbox)
                    if rerun_helper.test_connection():
                        rerun_all_folders = rerun_helper.mailbox.list_folders_recursive() or []
                        rerun_pids = set(processed_ids_by_mailbox.get(rerun_mailbox, []))
                        ext_reruns = self._scan_rerun_folder(
                            helper=rerun_helper, mailbox=rerun_mailbox,
                            all_folders=rerun_all_folders,
                            rerun_folder_name=rerun_folder_name,
                            recipient=recipient,
                            download_root=download_root, config=config,
                            mailbox_processed_ids=rerun_pids,
                            mark_as_processed=mark_as_processed,
                        )
                        if ext_reruns > 0:
                            logger.info(f"🔄 Handled {ext_reruns} RERUN(s) from {rerun_mailbox}")
                            processed_ids_by_mailbox[rerun_mailbox] = list(rerun_pids)[-5000:]
                            state["processed_ids_by_mailbox"] = processed_ids_by_mailbox
                            _save_json(self.state_path, state)
                    else:
                        logger.warning(f"⚠️ RERUN mailbox connection failed: {rerun_mailbox}")
                except Exception as e:
                    logger.error(f"❌ RERUN separate mailbox error: {e}")

            for mailbox in shared_mailboxes:
                helper = GraphAPIHelper(shared_mailbox=mailbox)
                if not helper.test_connection():
                    had_errors = True
                    self._status(f"שגיאת התחברות לתיבה: {mailbox}")
                    continue

                mailbox_processed_ids = set(processed_ids_by_mailbox.get(mailbox, []))
                if not mailbox_processed_ids and legacy_processed_ids:
                    mailbox_processed_ids = set(legacy_processed_ids)

                mailbox_last_checked = last_checked_by_mailbox.get(mailbox) or legacy_last_checked

                folder_id = None
                folder_name_lower = str(folder_name or "").strip().lower()

                all_folders = helper.mailbox.list_folders_recursive() or []
                for folder in all_folders:
                    folder_path = str(folder.get("path", "")).strip().lower()
                    folder_display = str(folder.get("displayName", "")).strip().lower()
                    if folder_path == folder_name_lower or folder_display == folder_name_lower:
                        folder_id = folder.get("id")
                        break

                if not folder_id:
                    folders = helper.mailbox.list_folders(show_inbox_only=False)
                    for folder in folders:
                        if folder.get("displayName", "").lower() == folder_name.lower():
                            folder_id = folder.get("id")
                            break
                if not folder_id:
                    inbox_folders = helper.mailbox.list_folders(show_inbox_only=True)
                    for folder in inbox_folders:
                        if folder.get("displayName", "").lower() == folder_name.lower():
                            folder_id = folder.get("id")
                            break

                # Fallback: try Graph API well-known folder name
                # (handles Hebrew "תיבת דואר נכנס" mapped to "Inbox", etc.)
                if not folder_id:
                    wkn_folder = helper.mailbox.get_well_known_folder(folder_name)
                    if wkn_folder and wkn_folder.get("id"):
                        folder_id = wkn_folder["id"]
                        logger.info(f"📂 Resolved '{folder_name}' via well-known name → {wkn_folder.get('displayName', '')}")

                if not folder_id:
                    had_errors = True
                    self._status(f"תיקייה לא נמצאה בתיבה {mailbox}: {folder_name}")
                    continue

                # ═══ RERUN FOLDER SCAN (START OF CYCLE) ═══
                start_reruns = 0
                if rerun_folder_name:
                    # If rerun_mailbox is set and differs from current mailbox,
                    # only run RERUN scan on the first mailbox iteration
                    if rerun_mailbox and rerun_mailbox.lower() != mailbox.lower():
                        # RERUN is in a different mailbox — handle separately below
                        pass
                    else:
                        start_reruns = self._scan_rerun_folder(
                            helper=helper, mailbox=mailbox,
                            all_folders=all_folders,
                            rerun_folder_name=rerun_folder_name,
                            recipient=recipient,
                            download_root=download_root, config=config,
                            mailbox_processed_ids=mailbox_processed_ids,
                            mark_as_processed=mark_as_processed,
                        )
                    if start_reruns > 0:
                        logger.info(f"🔄 Handled {start_reruns} RERUN(s) at cycle start")
                        # Persist immediately so IDs survive a stop
                        processed_ids_by_mailbox[mailbox] = list(mailbox_processed_ids)[-5000:]
                        state["processed_ids_by_mailbox"] = processed_ids_by_mailbox
                        _save_json(self.state_path, state)

                # ── Date filter for list_messages ──
                if heavy_only:
                    # Heavy-only mode: ignore last_checked — heavy emails may have
                    # been received long before the last regular scan.  Use
                    # scan_from_date so the full configured window is searched.
                    received_after = scan_from_datetime  # may be None → fetch all
                    logger.info(f"🏋️ Heavy-only date filter: scan_from_date={scan_from_datetime} (ignoring last_checked)")
                else:
                    received_after = None
                    if mailbox_last_checked:
                        try:
                            received_after = datetime.fromisoformat(str(mailbox_last_checked).replace("Z", "+00:00"))
                        except Exception as e:
                            logger.debug(f"Handled: {e}")
                            received_after = None

                    # Safety overlap: subtract 3 minutes to catch emails with delayed
                    # Graph API indexing.  Already-processed IDs are filtered out below,
                    # so the overlap is harmless.
                    if received_after is not None:
                        received_after -= timedelta(minutes=3)

                    # scan_from_date overrides if it's newer than last_checked
                    # (or if last_checked is None)
                    if scan_from_datetime:
                        if received_after is None or scan_from_datetime > received_after.replace(tzinfo=None):
                            received_after = scan_from_datetime
                            logger.info(f"📅 Using scan_from_date as filter: {scan_from_datetime}")

                # Record the moment we call list_messages — this becomes last_checked
                # to avoid skipping emails that arrive during the (potentially hours-long) processing
                fetch_timestamp = _now_iso()
                messages = helper.mailbox.list_messages(folder_id, limit=limit, received_after=received_after)
                logger.info(f"📬 list_messages returned {len(messages)} messages "
                            f"(received_after={received_after}, fetch_ts={fetch_timestamp})")
                messages.sort(key=lambda m: (
                    0 if (m.get("from", {}).get("emailAddress", {}).get("address", "") or "").lower() == mailbox.lower() else 1,
                    m.get("receivedDateTime", "")
                ))

                new_messages = [m for m in messages if m.get("id") and m.get("id") not in mailbox_processed_ids]
                logger.info(f"📬 {len(new_messages)} new messages after filtering "
                            f"{len(mailbox_processed_ids)} processed IDs")
                if not new_messages:
                    # Retry without last_checked filter, but STILL respect scan_from_date if configured
                    retry_after = scan_from_datetime if scan_from_datetime else None
                    if retry_after:
                        self._status(f"בודק שוב עם scan_from_date בלבד... ({mailbox})")
                    else:
                        self._status(f"בודק שוב ללא סינון תאריך... ({mailbox})")
                    messages = helper.mailbox.list_messages(folder_id, limit=limit, received_after=retry_after)
                    messages.sort(key=lambda m: (
                        0 if (m.get("from", {}).get("emailAddress", {}).get("address", "") or "").lower() == mailbox.lower() else 1,
                        m.get("receivedDateTime", "")
                    ))
                    new_messages = [m for m in messages if m.get("id") and m.get("id") not in mailbox_processed_ids]

                if not new_messages:
                    self._status(f"אין מיילים חדשים בתיבה: {mailbox}")
                    # Don't continue — still need to check RERUN folder below

                total_new_messages += len(new_messages)

                # ── Chronological sorting: oldest first ──
                if len(new_messages) > 1:
                    new_messages.sort(key=lambda m: m.get("receivedDateTime", ""))
                    logger.info(f"📊 Sorted {len(new_messages)} emails chronologically (oldest first)")

                # ── Heavy-only mode: keep only AI HEAVY emails ──
                if heavy_only:
                    new_messages = [
                        m for m in new_messages
                        if heavy_category_name in (m.get("categories") or [])
                    ]
                    # Retry with no date filter if no heavy emails found
                    if not new_messages:
                        logger.info(f"🏋️ No heavy emails in initial fetch — retrying without date filter")
                        retry_msgs = helper.mailbox.list_messages(folder_id, limit=limit, received_after=None)
                        retry_new = [m for m in retry_msgs if m.get("id") and m.get("id") not in mailbox_processed_ids]
                        new_messages = [
                            m for m in retry_new
                            if heavy_category_name in (m.get("categories") or [])
                        ]
                    if not new_messages:
                        self._status(f"אין מיילים כבדים ({heavy_category_name}) בתיבה: {mailbox}")
                    else:
                        logger.info(f"🏋️ Heavy-only mode: {len(new_messages)} emails with '{heavy_category_name}'")
                        self._status(f"🏋️ נמצאו {len(new_messages)} מיילים כבדים לעיבוד ({mailbox})")

                # Fetch category color map once per mailbox (for category banners in emails)
                _mailbox_category_colors = {}
                try:
                    _mailbox_category_colors = helper.mailbox.list_categories()
                except Exception:
                    pass

                all_messages_processed = True  # Track if loop completes without interruption

                for msg_idx, msg in enumerate(new_messages, 1):
                    # Check stop between emails
                    if self._stop_event.is_set():
                        logger.info(f"⏹️ Stop requested — finishing after {msg_idx-1}/{len(new_messages)} emails")
                        self._status(f"נעצר אחרי {msg_idx-1} מיילים (שמירת מצב...)")
                        all_messages_processed = False
                        break
                    
                    msg_id = msg.get("id")
                    if not msg_id:
                        continue

                    # ═══ Re-check state from disk (defense against stale in-memory set) ═══
                    try:
                        _fresh_state = _load_json(self.state_path, {})
                        _fresh_ids = set(_fresh_state.get("processed_ids_by_mailbox", {}).get(mailbox, []))
                        if msg_id in _fresh_ids and msg_id not in mailbox_processed_ids:
                            logger.info(f"⏭️ SKIP: msg {msg_id[:8]} was processed by another runner — skipping")
                            mailbox_processed_ids.add(msg_id)
                            continue
                    except Exception:
                        pass

                    # ═══ SKIP SENDERS CHECK ═══
                    msg_sender = (msg.get("from", {}).get("emailAddress", {}).get("address", "") or "").lower().strip()
                    if skip_senders and msg_sender in skip_senders:
                        logger.info(f"🚫 SKIP: mail {msg_idx}/{len(new_messages)} from {msg_sender} — in skip list")
                        self._status(f"🚫 דילוג על מייל {msg_idx}/{len(new_messages)} — שולח ברשימת דילוג")
                        _append_log(log_path, {
                            "event": "skip_sender",
                            "message_id": msg_id,
                            "mailbox": mailbox,
                            "sender": msg_sender,
                            "timestamp": _now_iso()
                        })
                        if mark_as_processed:
                            try:
                                helper.ensure_category("AI SKIPPED", "preset8")
                                helper.mark_message_processed(msg_id, "AI SKIPPED")
                            except Exception:
                                pass
                        mailbox_processed_ids.add(msg_id)
                        continue
                    # ═══ END SKIP SENDERS ═══

                    # ═══ SKIP BY CATEGORY ═══
                    msg_categories_list = msg.get("categories") or []
                    if skip_categories and msg_categories_list:
                        matched_cats = skip_categories & set(msg_categories_list)
                        if matched_cats:
                            logger.info(f"🏷️ SKIP: mail {msg_idx}/{len(new_messages)} — categories {matched_cats} in skip list")
                            self._status(f"🏷️ דילוג על מייל {msg_idx}/{len(new_messages)} — קטגוריה ברשימת דילוג")
                            _append_log(log_path, {
                                "event": "skip_category",
                                "message_id": msg_id,
                                "mailbox": mailbox,
                                "categories": sorted(matched_cats),
                                "timestamp": _now_iso()
                            })
                            mailbox_processed_ids.add(msg_id)
                            continue
                    # ═══ END SKIP BY CATEGORY ═══

                    # ═══ SKIP AI HEAVY (already marked — waiting for heavy run) ═══
                    if not heavy_only and heavy_category_name in msg_categories_list:
                        logger.info(f"🏷️ SKIP: mail {msg_idx}/{len(new_messages)} already marked '{heavy_category_name}' — waiting for heavy run")
                        continue
                    # ═══ END SKIP AI HEAVY ═══

                    # ═══ RERUN DETECTION (disabled — use RERUN folder only) ═══
                    # Previously: sender==mailbox → treat as RERUN.
                    # Disabled because mails can appear in the inbox from
                    # forwards / rules / bounces.  RERUN is now handled
                    # exclusively via the dedicated RERUN folder.
                    is_rerun = msg_sender == mailbox.lower().strip()
                    if is_rerun:
                        logger.info(f"⏭️ SKIP: sender is ourselves ({msg_sender}) — ignoring in inbox (use RERUN folder)")
                        mailbox_processed_ids.add(msg_id)
                        if mark_as_processed:
                            try:
                                helper.ensure_category("AI RERUN", "preset9")
                                helper.mark_message_processed(msg_id, "AI RERUN")
                            except Exception:
                                pass
                        continue  # Skip — not a valid incoming email
                    # ═══ END RERUN ═══

                    # ═══ SKIP NO ATTACHMENTS (pre-download) ═══
                    if msg.get("hasAttachments") is False:
                        logger.info(f"📭 SKIP: mail {msg_idx}/{len(new_messages)} has no attachments — marking NO DRAW")
                        self._status(f"📭 דילוג על מייל {msg_idx}/{len(new_messages)} — ללא קבצים מצורפים")
                        _append_log(log_path, {
                            "event": "no_attachment_skipped",
                            "message_id": msg_id,
                            "mailbox": mailbox,
                            "timestamp": _now_iso()
                        })
                        if mark_as_processed:
                            try:
                                helper.ensure_category(nodraw_category_name, nodraw_category_color)
                                helper.mark_message_processed(msg_id, nodraw_category_name)
                                logger.info(f"📭 Marked as '{nodraw_category_name}' (no attachments)")
                            except Exception:
                                pass
                        mailbox_processed_ids.add(msg_id)
                        continue
                    # ═══ END SKIP NO ATTACHMENTS ═══

                    run_stamp = datetime.now().strftime("%Y%m%d%H%M%S")
                    mailbox_tag = mailbox.replace("@", "_at_").replace(".", "_")
                    run_dir = download_root / f"auto_{run_stamp}_{mailbox_tag}_{msg_id[:8]}"
                    self._status(f"מוריד מייל {msg_idx}/{len(new_messages)} | {mailbox}: {msg_id[:8]}")
                    download_result = helper.download_message_by_id(msg_id, str(run_dir))

                    if not download_result.get("success"):
                        logger.warning(f"⚠️ Download FAILED for mail {msg_idx}/{len(new_messages)} | {mailbox}: {msg_id[:8]} — skipping")
                        self._status(f"שגיאה בהורדת מייל {msg_idx}/{len(new_messages)} | {mailbox}: {msg_id[:8]}")
                        continue

                    message_dir = Path(download_result.get("message_dir"))
                    sender = download_result.get("sender", "").strip()
                    original_body_html = download_result.get("original_body_html", "")
                    original_body_type = download_result.get("original_body_type", "Text")
                    source_web_link = download_result.get("web_link", "")
                    msg_categories = download_result.get("categories", [])

                    # ── Heavy email check: skip if too many drawing files (not in heavy_only mode) ──
                    if not heavy_only and max_files_per_email > 0:
                        _drawing_count = _count_drawing_files(message_dir)
                        if _drawing_count > max_files_per_email:
                            logger.warning(
                                f"⚠️ AI HEAVY: {_drawing_count} drawing files in msg {msg_id[:8]} "
                                f"(threshold={max_files_per_email}) — skipping"
                            )
                            self._status(f"⚠️ מייל כבד ({_drawing_count} קבצים) — מדלג {msg_id[:8]}")
                            if mark_as_processed:
                                try:
                                    helper.ensure_category(heavy_category_name, heavy_category_color)
                                    mark_ok = helper.mark_message_processed(msg_id, heavy_category_name)
                                    if mark_ok:
                                        logger.info(f"🏷️ Marked as '{heavy_category_name}' ({_drawing_count} files)")
                                    else:
                                        logger.warning(f"⚠️ mark_message_processed FAILED for '{heavy_category_name}' — msg {msg_id[:20]}…")
                                except Exception as e:
                                    logger.warning(f"⚠️ Exception marking heavy mail: {e}")
                            if cleanup_download:
                                try:
                                    if Path(run_dir).exists():
                                        shutil.rmtree(run_dir)
                                except Exception:
                                    pass
                            _append_log(log_path, {
                                "event": "heavy_email_skipped",
                                "message_id": msg_id,
                                "mailbox": mailbox,
                                "drawing_files": _drawing_count,
                                "threshold": max_files_per_email,
                                "timestamp": _now_iso()
                            })
                            continue

                    # ═══ SKIP NO DRAWINGS (post-download, pre-AI) ═══
                    _drawing_count_check = _count_drawing_files(message_dir)
                    if _drawing_count_check == 0:
                        logger.info(f"📭 NO DRAW: 0 drawing files in {msg_id[:8]} — skipping AI pipeline")
                        self._status(f"📭 אין שרטוטים במייל {msg_idx}/{len(new_messages)} — מדלג")
                        if mark_as_processed:
                            try:
                                helper.ensure_category(nodraw_category_name, nodraw_category_color)
                                helper.mark_message_processed(msg_id, nodraw_category_name)
                                logger.info(f"📭 Marked as '{nodraw_category_name}' (no drawings)")
                            except Exception:
                                pass
                        if cleanup_download:
                            try:
                                if Path(run_dir).exists():
                                    shutil.rmtree(run_dir)
                            except Exception:
                                pass
                        mailbox_processed_ids.add(msg_id)
                        _append_log(log_path, {
                            "event": "no_draw_skipped",
                            "message_id": msg_id,
                            "mailbox": mailbox,
                            "timestamp": _now_iso()
                        })
                        continue
                    # ═══ END SKIP NO DRAWINGS ═══

                    self._status(f"מריץ ניתוח [{msg_idx}/{len(new_messages)}] | {mailbox}...")
                    _scan_start = time.time()
                    try:
                        results, output_folder, output_path, cost_summary, file_classifications = _scan_folder_compat(
                            message_dir=message_dir,
                            recursive=recursive,
                            stages=stages,
                            enable_retry=enable_retry,
                            tosend_folder=tosend_folder,
                            confidence_level=confidence_level,
                            stage1_skip_retry_resolution_px=stage1_skip_retry_resolution_px,
                            max_file_size_mb=max_file_size_mb,
                            max_image_dimension=max_image_dimension,
                        )
                    except Exception as process_error:
                        logger.error(f"❌ scan_folder FAILED for mail {msg_id[:8]}: {process_error}", exc_info=True)
                        had_errors = True
                        self._notify_failure(f"Mail {msg_id[:8]} failed: {str(process_error)[:200]}")
                        self._status(f"שגיאה בעיבוד מייל {msg_idx}/{len(new_messages)} | {mailbox}: {msg_id[:8]} - ממשיך")
                        continue
                    _scan_elapsed = time.time() - _scan_start

                    # ── Detect NO DRAW: no actual drawings processed ──
                    _scan_files_processed = 0
                    if isinstance(cost_summary, dict):
                        _scan_files_processed = cost_summary.get('successful_files', 0) or cost_summary.get('total_files', 0) or 0
                    is_no_draw = (_scan_files_processed == 0 and (not results or len(results) == 0))
                    if is_no_draw:
                        logger.warning(f"⚠️ NO DRAW: no drawings found for message {msg_id[:8]} — will not send email")

                    request_id = None
                    try:
                        b2b_files = list(Path(message_dir).rglob("B2B-*.txt"))
                        if b2b_files:
                            b2b_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                            b2b_filename = b2b_files[0].stem
                            request_id = b2b_filename.split("-", 2)[-1]
                    except Exception as e:
                        logger.error(f"Failed to extract request_id: {e}")

                    if not request_id and output_path and Path(output_path).exists():
                        try:
                            request_id = Path(output_path).stem.replace("SUMMARY_all_results_", "")
                        except Exception as e:
                            logger.debug(f"Handled: {e}")
                            request_id = run_stamp

                    if not request_id:
                        request_id = run_stamp

                    attachments = []
                    if output_path and Path(output_path).exists():
                        attachments.append(Path(output_path))

                        try:
                            timestamp = Path(output_path).stem.replace("SUMMARY_all_results_", "")
                            classification = Path(output_folder) / f"SUMMARY_all_classifications_{timestamp}.xlsx"
                            if classification.exists():
                                attachments.append(classification)
                        except Exception as e:
                            logger.debug(f"Ignored: {e}")
                            pass

                    if output_copy_folder and attachments:
                        try:
                            copy_root = Path(output_copy_folder) / run_stamp
                            copy_root.mkdir(parents=True, exist_ok=True)
                            for file_path in attachments:
                                if str(file_path).lower().endswith('.zip'):
                                    continue
                                (copy_root / file_path.name).write_bytes(Path(file_path).read_bytes())
                        except Exception as e:
                            logger.debug(f"Ignored: {e}")
                            pass

                    tosend_dest = None
                    if tosend_folder:
                        tosend_root = Path(tosend_folder)
                        # file_utils creates TO_SEND using subfolder names, not message_dir name
                        # Find the most recently created *_TO_SEND folder in tosend_root
                        # that was created during this run (within the last few minutes)
                        import time as _time_mod
                        _now = _time_mod.time()
                        _to_send_candidates = []
                        for d in tosend_root.iterdir():
                            if d.is_dir() and d.name.endswith("_TO_SEND"):
                                age = _now - d.stat().st_mtime
                                if age < 300:  # created in last 5 minutes
                                    _to_send_candidates.append((d, age))
                        if _to_send_candidates:
                            _to_send_candidates.sort(key=lambda x: x[1])  # newest first
                            tosend_dest = _to_send_candidates[0][0]
                            logger.info(f"Found TO_SEND folder: {tosend_dest.name}")
                        else:
                            # Fallback: try old naming convention
                            tosend_dest = tosend_root / f"{message_dir.name}_TO_SEND"

                    display_names_sent = []
                    if auto_send and not is_no_draw:
                        all_files_to_send = []
                        display_name_map = {}
                        if file_classifications:
                            for fc in file_classifications:
                                display_name = fc.get('display_name', '').strip()
                                renamed_filename = fc.get('renamed_filename', '').strip()
                                if display_name and renamed_filename:
                                    display_name_map[renamed_filename] = display_name

                        if tosend_dest and tosend_dest.exists():
                            for file_path in tosend_dest.glob("*"):
                                if file_path.is_file() and not file_path.suffix.lower() == '.zip':
                                    display_name = display_name_map.get(file_path.name, '')
                                    if display_name:
                                        all_files_to_send.append({
                                            'path': file_path,
                                            'display_name': display_name
                                        })
                                        logger.info(f"FILE WITH DISPLAY NAME '{display_name}' WAS SENT")
                                    else:
                                        all_files_to_send.append(file_path)

                        if all_files_to_send:
                            email_body_html = original_body_html or ""

                            # קרא webLink מ-email.txt אם אין לנו אותו
                            if not source_web_link and message_dir:
                                _email_txt = message_dir / "email.txt"
                                if _email_txt.exists():
                                    try:
                                        with open(_email_txt, "r", encoding="utf-8") as _ef:
                                            for _line in _ef:
                                                if _line.startswith("מייל מקור: "):
                                                    source_web_link = _line.replace("מייל מקור: ", "").strip()
                                                    break
                                    except Exception:
                                        pass

                            if output_path and Path(output_path).exists():
                                try:
                                    import pandas as pd
                                    df = pd.read_excel(output_path)

                                    display_columns = ['file_name', 'customer_name', 'merged_description', 'quantity', 'merged_bom', 'merged_notes', 'part_number', 'revision', 'drawing_number']
                                    html_table = '<table border="1" cellpadding="8" cellspacing="0" style="width:100%; border-collapse:collapse; direction: rtl; text-align: right;">'

                                    header_mapping = {
                                        'file_name': 'שם קובץ',
                                        'customer_name': 'שם לקוח',
                                        'merged_description': 'תיאור מורחב',
                                        'quantity': 'כמות',
                                        'merged_bom': 'עץ מוצר',
                                        'merged_notes': 'הערות מיוחדות של לקוח',
                                        'part_number': 'מספר פריט',
                                        'revision': 'מספר גירסא',
                                        'drawing_number': 'מספר שרטוט'
                                    }

                                    html_table += '<tr style="background-color: #4472C4; color: white; font-weight: bold;">\n'
                                    for col in display_columns:
                                        header_name = header_mapping.get(col, col)
                                        html_table += f'<th>{header_name}</th>\n'
                                    html_table += '</tr>\n'

                                    df = df.fillna('')
                                    ltr_columns = {'part_number', 'revision', 'drawing_number'}

                                    # Ensure merged/summary columns exist
                                    for col in ('merged_description', 'merged_bom', 'merged_notes', 'quantity'):
                                        if col not in df.columns:
                                            df[col] = ''
                                        else:
                                            df[col] = df[col].astype(str).str.strip()
                                            df.loc[df[col] == 'nan', col] = ''

                                    has_pl_override_col = 'part_number_ocr_original' in df.columns

                                    for _, row in df.iterrows():
                                        confidence = str(row.get('confidence_level', '')).strip().upper()

                                        if confidence in ['HIGH', 'FULL']:
                                            row_color = '#C6E0B4'
                                        elif confidence == 'MEDIUM':
                                            row_color = '#FFE699'
                                        else:
                                            row_color = '#F4B084'

                                        is_pl_override = False
                                        is_pl_confirmed = False
                                        if has_pl_override_col:
                                            ocr_orig = str(row.get('part_number_ocr_original', '')).strip()
                                            if ocr_orig and ocr_orig != 'nan':
                                                is_pl_override = True
                                        # Check for AS PL (PL confirmed, no override)
                                        if 'pl_override_note' in df.columns:
                                            pl_note = str(row.get('pl_override_note', '')).strip()
                                            if pl_note == 'AS PL':
                                                is_pl_confirmed = True

                                        html_table += f'<tr style="background-color: {row_color};">\n'
                                        for col in display_columns:
                                            value = row.get(col, "")
                                            if col == 'part_number' and is_pl_override:
                                                html_table += (
                                                    '<td dir="ltr" style="text-align:left; background-color: #DAEEF3; unicode-bidi:bidi-override;">'
                                                    f'<font color="#00008B"><b><span dir="ltr" style="unicode-bidi:bidi-override;">&lrm;{value}</span></b></font>'
                                                    ' <font size="1" color="#00008B">(PL)</font>'
                                                    '</td>\n'
                                                )
                                            elif col == 'part_number' and is_pl_confirmed:
                                                html_table += (
                                                    '<td dir="ltr" style="text-align:left; background-color: #E2EFDA; unicode-bidi:bidi-override;">'
                                                    f'<span dir="ltr" style="unicode-bidi:bidi-override;">&lrm;{value}</span>'
                                                    ' <font size="1" color="#2E7D32">(AS PL)</font>'
                                                    '</td>\n'
                                                )
                                            elif col in ltr_columns:
                                                value = f'&lrm;{value}'
                                                html_table += (
                                                    '<td dir="ltr" style="text-align:left; unicode-bidi:bidi-override;">'
                                                    f'<span dir="ltr" style="unicode-bidi:bidi-override;">{value}</span>'
                                                    '</td>\n'
                                                )
                                            elif col in ('merged_bom', 'merged_description'):
                                                # Multi-line fields: newlines → <br>, alternates → italic
                                                cell = str(value or '')
                                                # Italicise individual (חלופי) items
                                                cell = re.sub(
                                                    r'([^|,<\n]+\(חלופי\)[^|,<\n]*)',
                                                    r'<i>\1</i>',
                                                    cell,
                                                )
                                                cell = cell.replace('\n', '<br>')
                                                html_table += f'<td>{cell}</td>\n'
                                            else:
                                                html_table += f'<td>{value}</td>\n'
                                        html_table += '</tr>\n'

                                    html_table += '</table>\n<br>\n'

                                    # הוסף לינק למייל מקור
                                    source_link_html = ""
                                    if source_web_link:
                                        source_link_html = (
                                            '<p style="margin-top:10px;">'
                                            f'<a href="{source_web_link}" style="color:#0078D4;text-decoration:none;font-size:13px;">'
                                            '\U0001f4e7 פתח מייל מקור ב-Outlook</a></p>'
                                        )

                                    email_body_html = html_table + source_link_html + email_body_html

                                except Exception as e:
                                    logger.debug(f"Error: {e}")
                                    import traceback
                                    traceback.print_exc()

                            # Fallback: הוסף לינק למייל מקור גם אם לא נבנתה טבלה
                            if source_web_link and 'פתח מייל מקור' not in email_body_html:
                                source_link_html = (
                                    '<p style="margin-top:10px;">'
                                    f'<a href="{source_web_link}" style="color:#0078D4;text-decoration:none;font-size:13px;">'
                                    '\U0001f4e7 פתח מייל מקור ב-Outlook</a></p>'
                                )
                                email_body_html = source_link_html + email_body_html

                            # Add category banner from original email
                            if msg_categories:
                                cat_banner = _build_category_banner(msg_categories, _mailbox_category_colors)
                                email_body_html = cat_banner + email_body_html

                            email_subject = f"B2B_Quotation_{request_id or run_stamp}"
                            if sender:
                                email_subject = f"{email_subject} | {sender}"

                            self._status(f"שולח מייל... ({mailbox})")
                            for att in all_files_to_send:
                                if isinstance(att, dict) and att.get('display_name'):
                                    display_name = att.get('display_name')
                                    display_names_sent.append(display_name)
                                    logger.info(f"FILE WITH DISPLAY NAME '{display_name}' WAS SENT")
                            helper.send_email(
                                to_address=recipient,
                                subject=email_subject,
                                body=email_body_html,
                                attachments=all_files_to_send,
                                body_type="HTML",
                                replace_display_with_filename=True
                            )

                    _cost_usd = 0.0
                    _files_processed = 0
                    if isinstance(cost_summary, dict):
                        _cost_usd = cost_summary.get('total_cost', 0) or 0
                        _files_processed = cost_summary.get('successful_files', 0) or cost_summary.get('total_files', 0) or 0

                    _accuracy_data = {"full": 0, "high": 0, "medium": 0, "low": 0, "none": 0, "total": 0}
                    _customers = []
                    _pl_overrides = 0
                    _error_types = []
                    try:
                        _acc_path = output_path if output_path and Path(output_path).exists() else None
                        if _acc_path:
                            import pandas as pd
                            df_acc = pd.read_excel(_acc_path)
                            if 'confidence_level' in df_acc.columns:
                                for _, _acc_row in df_acc.iterrows():
                                    conf = str(_acc_row.get('confidence_level', '')).strip().lower()
                                    pn = str(_acc_row.get('part_number', '')).strip()
                                    _accuracy_data["total"] += 1
                                    if not pn or pn == 'nan' or pn == '':
                                        _accuracy_data["none"] += 1
                                        _error_types.append("missing_part_number")
                                    elif conf in ('full', 'high', 'medium', 'low'):
                                        _accuracy_data[conf] += 1
                                        if conf == 'low':
                                            _error_types.append("low_confidence")
                                    else:
                                        _accuracy_data["none"] += 1

                                    cust = str(_acc_row.get('customer_name', '')).strip()
                                    if cust and cust != 'nan':
                                        _customers.append(cust)

                                    ocr_orig = str(_acc_row.get('part_number_ocr_original', '')).strip()
                                    if ocr_orig and ocr_orig != 'nan':
                                        _pl_overrides += 1
                    except Exception as e:
                        logger.debug(f"Ignored: {e}")
                        pass

                    _append_log(log_path, {
                        "id": str(uuid4()),
                        "shared_mailbox": mailbox,
                        "message_id": msg_id,
                        "run_type": "heavy" if heavy_only else "regular",
                        "received": download_result.get("received"),
                        "sender": sender,
                        "download_dir": str(run_dir),
                        "tosend": str(tosend_folder) if tosend_folder else "",
                        "output_summary": str(output_path) if output_path else "",
                        "sent": auto_send,
                        "attachments_display_names": display_names_sent if auto_send else [],
                        "processing_time_seconds": round(_scan_elapsed, 1),
                        "processing_time_minutes": round(_scan_elapsed / 60, 2),
                        "cost_usd": round(float(_cost_usd), 4),
                        "files_processed": int(_files_processed),
                        "accuracy_data": _accuracy_data,
                        "customers": list(set(_customers)),
                        "pl_overrides": _pl_overrides,
                        "error_types": _error_types,
                        "items_count": _accuracy_data["total"],
                        "timestamp": _now_iso()
                    })

                    if mark_as_processed:
                        try:
                            if is_no_draw:
                                # No drawings found → mark with NO DRAW category
                                self._status(f"מסמן מייל NO DRAW... ({mailbox})")
                                helper.ensure_category(nodraw_category_name, nodraw_category_color)
                                if heavy_only:
                                    mark_ok = helper.replace_category(msg_id, heavy_category_name, nodraw_category_name)
                                else:
                                    mark_ok = helper.mark_message_processed(msg_id, nodraw_category_name)
                                if mark_ok:
                                    logger.info(f"📭 Marked as '{nodraw_category_name}' (no drawings)")
                                else:
                                    logger.warning(f"⚠️ mark_message_processed FAILED for '{nodraw_category_name}' — msg {msg_id[:20]}…  error={helper.last_error}")
                            else:
                                self._status(f"מסמן מייל כמעובד... ({mailbox})")
                                helper.ensure_category(mark_category_name, mark_category_color)
                                if heavy_only:
                                    mark_ok = helper.replace_category(msg_id, heavy_category_name, mark_category_name)
                                else:
                                    mark_ok = helper.mark_message_processed(msg_id, mark_category_name)
                                if mark_ok:
                                    logger.info(f"✅ Marked as '{mark_category_name}'")
                                else:
                                    logger.warning(f"⚠️ mark_message_processed FAILED for '{mark_category_name}' — msg {msg_id[:20]}…  error={helper.last_error}")
                        except Exception as e:
                            logger.warning(f"⚠️ Exception while marking mail ({mailbox}): {e}")
                            self._status(f"⚠️ לא ניתן לסמן מייל ({mailbox}): {e}")

                    if cleanup_download:
                        try:
                            if Path(run_dir).exists():
                                shutil.rmtree(run_dir)
                        except Exception as e:
                            logger.debug(f"Ignored: {e}")
                            pass

                    mailbox_processed_ids.add(msg_id)
                    # Update the dict BEFORE building state — otherwise per-message save writes empty state
                    processed_ids_by_mailbox[mailbox] = mailbox_processed_ids
                    state["processed_ids_by_mailbox"] = {
                        k: list(v)[-5000:] if isinstance(v, set) else v[-5000:]
                        for k, v in processed_ids_by_mailbox.items()
                    }
                    state["last_checked_by_mailbox"] = last_checked_by_mailbox
                    combined = set()
                    for ids_list in processed_ids_by_mailbox.values():
                        combined.update(ids_list if isinstance(ids_list, set) else set(ids_list))
                    state["processed_ids"] = list(combined)[-5000:]
                    _save_json(self.state_path, state)
                    logger.info(f"\U0001f4e7 Mail processed & saved: {msg_id[:20]}... (total: {len(mailbox_processed_ids)})")

                    # ─── RERUN PRIORITY CHECK (inbox sender-match disabled) ───
                    # Inbox-based sender==mailbox RERUN removed — only RERUN folder is used.

                    # Check dedicated RERUN folder between emails
                    if rerun_folder_name:
                        try:
                            # Use separate helper if RERUN is in a different mailbox
                            if rerun_mailbox and rerun_mailbox.lower() != mailbox.lower():
                                rerun_helper_mid = GraphAPIHelper(shared_mailbox=rerun_mailbox)
                                rerun_folders_mid = rerun_helper_mid.mailbox.list_folders_recursive() or []
                                rerun_pids_mid = set(processed_ids_by_mailbox.get(rerun_mailbox, []))
                                folder_reruns = self._scan_rerun_folder(
                                    helper=rerun_helper_mid, mailbox=rerun_mailbox,
                                    all_folders=rerun_folders_mid,
                                    rerun_folder_name=rerun_folder_name,
                                    recipient=recipient,
                                    download_root=download_root, config=config,
                                    mailbox_processed_ids=rerun_pids_mid,
                                    mark_as_processed=mark_as_processed,
                                )
                                if folder_reruns > 0:
                                    logger.info(f"🔄 Handled {folder_reruns} RERUN(s) between emails (separate mailbox)")
                                    processed_ids_by_mailbox[rerun_mailbox] = list(rerun_pids_mid)[-5000:]
                                    state["processed_ids_by_mailbox"] = processed_ids_by_mailbox
                                    _save_json(self.state_path, state)
                            else:
                                folder_reruns = self._scan_rerun_folder(
                                    helper=helper, mailbox=mailbox,
                                    all_folders=all_folders,
                                    rerun_folder_name=rerun_folder_name,
                                    recipient=recipient,
                                    download_root=download_root, config=config,
                                    mailbox_processed_ids=mailbox_processed_ids,
                                    mark_as_processed=mark_as_processed,
                                )
                                if folder_reruns > 0:
                                    logger.info(f"🔄 Handled {folder_reruns} RERUN(s) between emails (folder)")
                                    processed_ids_by_mailbox[mailbox] = list(mailbox_processed_ids)[-5000:]
                                    state["processed_ids_by_mailbox"] = processed_ids_by_mailbox
                                    _save_json(self.state_path, state)
                        except Exception as e:
                            logger.debug(f"RERUN folder check error (non-critical): {e}")

                # ═══ RERUN FOLDER SCAN (END OF CYCLE) ═══
                if rerun_folder_name:
                    if rerun_mailbox and rerun_mailbox.lower() != mailbox.lower():
                        # Use separate helper for different RERUN mailbox
                        try:
                            rerun_helper_end = GraphAPIHelper(shared_mailbox=rerun_mailbox)
                            rerun_folders_end = rerun_helper_end.mailbox.list_folders_recursive() or []
                            rerun_pids_end = set(processed_ids_by_mailbox.get(rerun_mailbox, []))
                            end_reruns = self._scan_rerun_folder(
                                helper=rerun_helper_end, mailbox=rerun_mailbox,
                                all_folders=rerun_folders_end,
                                rerun_folder_name=rerun_folder_name,
                                recipient=recipient,
                                download_root=download_root, config=config,
                                mailbox_processed_ids=rerun_pids_end,
                                mark_as_processed=mark_as_processed,
                            )
                            if end_reruns > 0:
                                logger.info(f"🔄 Handled {end_reruns} RERUN(s) at cycle end (separate mailbox)")
                                processed_ids_by_mailbox[rerun_mailbox] = list(rerun_pids_end)[-5000:]
                                state["processed_ids_by_mailbox"] = processed_ids_by_mailbox
                                _save_json(self.state_path, state)
                        except Exception as e:
                            logger.debug(f"RERUN end-of-cycle check error: {e}")
                    else:
                        end_reruns = self._scan_rerun_folder(
                            helper=helper, mailbox=mailbox,
                            all_folders=all_folders,
                            rerun_folder_name=rerun_folder_name,
                            recipient=recipient,
                            download_root=download_root, config=config,
                            mailbox_processed_ids=mailbox_processed_ids,
                            mark_as_processed=mark_as_processed,
                        )
                        if end_reruns > 0:
                            logger.info(f"🔄 Handled {end_reruns} RERUN(s) at cycle end")
                            processed_ids_by_mailbox[mailbox] = list(mailbox_processed_ids)[-5000:]
                            state["processed_ids_by_mailbox"] = processed_ids_by_mailbox
                            _save_json(self.state_path, state)

                processed_ids_by_mailbox[mailbox] = list(mailbox_processed_ids)[-5000:]
                # Use the fetch timestamp (when list_messages was called), NOT current time.
                # This prevents the gap: emails arriving during hours-long processing
                # would be behind _now_iso() but after fetch_timestamp.
                # IMPORTANT: Only advance last_checked if ALL messages were processed.
                # If loop was interrupted (stop event), keep old last_checked so
                # unprocessed messages will be retried on the next cycle.
                if heavy_only:
                    logger.info(f"🏋️ Heavy-only run — last_checked NOT advanced for {mailbox}")
                elif all_messages_processed:
                    last_checked_by_mailbox[mailbox] = fetch_timestamp
                else:
                    logger.info(f"⚠️ last_checked NOT advanced for {mailbox} — "
                                f"{len(new_messages)} messages were not all processed")
                    # Still update last_checked to the oldest unprocessed message's
                    # receivedDateTime minus 1 second, so we don't re-fetch already-processed
                    # messages but DO re-fetch unprocessed ones.
                    # processed_ids will filter out already-handled messages regardless.

            combined_processed_ids = set()
            for ids in processed_ids_by_mailbox.values():
                combined_processed_ids.update(ids)

            state["processed_ids_by_mailbox"] = processed_ids_by_mailbox
            state["last_checked_by_mailbox"] = last_checked_by_mailbox
            state["processed_ids"] = list(combined_processed_ids)[-5000:]
            state["last_checked"] = _now_iso()
            _save_json(self.state_path, state)

            # Calculate next run time
            if getattr(self, '_in_loop', False):
                next_run_time = datetime.now() + timedelta(minutes=config.get('poll_interval_minutes', 10))
                next_run_str = next_run_time.strftime('%H:%M:%S')
                if total_new_messages == 0:
                    self._status(f"אין מיילים חדשים בכל התיבות | הסבב הבא: {next_run_str}")
                elif had_errors:
                    self._status(f"הסבב הושלם עם שגיאות חלקיות | הסבב הבא: {next_run_str}")
                else:
                    self._status(f"הסבב הושלם | הסבב הבא: {next_run_str}")
            else:
                if total_new_messages == 0:
                    self._status("אין מיילים חדשים בכל התיבות | הסבב הסתיים")
                elif had_errors:
                    self._status("הסבב הושלם עם שגיאות חלקיות")
                else:
                    self._status("הסבב הושלם בהצלחה")

            # Export report at end of cycle (loop + one-shot), if configured.
            self._export_cycle_report_if_configured(config, heavy_only, cycle_start)

        except Exception as e:
            import traceback
            logger.error(f"\u274c Automation run FAILED: {e}\n{traceback.format_exc()}")
            self._notify_failure(f"Run failed: {str(e)[:200]}")
            self._status(f"שגיאה באוטומציה: {e}")
