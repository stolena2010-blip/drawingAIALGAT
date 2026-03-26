"""
Log Reader — Read and filter automation JSONL logs + live log files
"""
import json
import os
from pathlib import Path
from typing import List, Dict, Any, Optional
from datetime import datetime, timedelta, timezone

LOG_DIR = Path(__file__).resolve().parent.parent.parent
JSONL_PATH = LOG_DIR / "automation_log.jsonl"
LOGS_FOLDER = LOG_DIR / "logs"


def _read_jsonl_file(path: Path) -> List[Dict[str, Any]]:
    """Read all valid JSON entries from a single JSONL file."""
    entries = []
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    except Exception:
        pass
    return entries


def load_log_entries(max_entries: int = 10000) -> List[Dict[str, Any]]:
    """Load entries from ALL automation_log*.jsonl files, deduplicated by id.
    Matches original dashboard_gui.py behavior."""
    seen_ids: set = set()
    all_entries: List[Dict[str, Any]] = []

    # Gather all log files: main + backups + .bak + dated
    log_files = list(LOG_DIR.glob("automation_log*.jsonl"))
    bak = LOG_DIR / "automation_log.jsonl.bak"
    if bak.exists() and bak not in log_files:
        log_files.append(bak)

    for path in log_files:
        for entry in _read_jsonl_file(path):
            eid = entry.get("id", "")
            if eid and eid in seen_ids:
                continue
            if eid:
                seen_ids.add(eid)
            all_entries.append(entry)

    # Sort by timestamp ascending, then reverse for most-recent-first
    def _ts(e):
        ts = e.get("timestamp") or e.get("start_time") or ""
        try:
            return datetime.fromisoformat(ts.replace("Z", "+00:00")).replace(tzinfo=None)
        except (ValueError, TypeError):
            return datetime.min

    all_entries.sort(key=_ts, reverse=True)
    return all_entries[:max_entries]


def filter_by_period(entries: List[Dict[str, Any]], period: str,
                     date_from: str = "", date_to: str = "") -> List[Dict[str, Any]]:
    """Filter entries by period: today, week, month, all, or custom range."""
    if period == "הכל" and not date_from and not date_to:
        return entries
    if not entries:
        return entries

    now = datetime.now(timezone.utc).replace(tzinfo=None)
    cutoff_start = None
    cutoff_end = None

    if period == "היום":
        cutoff_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    elif period == "שבוע":
        cutoff_start = now - timedelta(days=7)
    elif period == "חודש":
        cutoff_start = now - timedelta(days=30)
    elif period == "טווח" and date_from:
        try:
            cutoff_start = datetime.strptime(date_from, "%Y-%m-%d")
        except ValueError:
            pass
        if date_to:
            try:
                cutoff_end = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)
            except ValueError:
                pass

    if cutoff_start is None and cutoff_end is None:
        return entries

    filtered = []
    for e in entries:
        ts = e.get("timestamp") or e.get("start_time") or ""
        try:
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).replace(tzinfo=None)
            if cutoff_start and dt < cutoff_start:
                continue
            if cutoff_end and dt >= cutoff_end:
                continue
            filtered.append(e)
        except (ValueError, AttributeError):
            filtered.append(e)
    return filtered


def get_accuracy_weights() -> Dict[str, float]:
    """Load accuracy weights from environment (same as original dashboard)."""
    return {
        "full":   float(os.getenv("ACCURACY_WEIGHT_FULL", "1.0")),
        "high":   float(os.getenv("ACCURACY_WEIGHT_HIGH", "1.0")),
        "medium": float(os.getenv("ACCURACY_WEIGHT_MEDIUM", "0.8")),
        "low":    float(os.getenv("ACCURACY_WEIGHT_LOW", "0.5")),
        "none":   float(os.getenv("ACCURACY_WEIGHT_NONE", "0.0")),
    }


def calc_weighted_accuracy(accuracy_data: Dict[str, Any], weights: Dict[str, float]) -> float:
    """Calculate weighted accuracy % for one entry. Returns 0-100."""
    total = int(accuracy_data.get("total", 0) or 0)
    if total == 0:
        return 0.0
    score = sum(
        int(accuracy_data.get(level, 0) or 0) * weights.get(level, 0)
        for level in ("full", "high", "medium", "low", "none")
    )
    return (score / total) * 100


def save_entry_field(entry_id: str, field: str, value: Any) -> bool:
    """Update a single field on a log entry in-place.
    Rewrites the JSONL file with the updated entry. Returns True on success."""
    if not JSONL_PATH.exists():
        return False
    lines = JSONL_PATH.read_text(encoding="utf-8").splitlines()
    updated = False
    new_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            new_lines.append(line)
            continue
        try:
            entry = json.loads(stripped)
        except json.JSONDecodeError:
            new_lines.append(line)
            continue
        if entry.get("id") == entry_id:
            entry[field] = value
            new_lines.append(json.dumps(entry, ensure_ascii=False))
            updated = True
        else:
            new_lines.append(line)
    if updated:
        JSONL_PATH.write_text("\n".join(new_lines) + "\n", encoding="utf-8")
    return updated


def get_latest_log_file() -> Optional[Path]:
    """Find the most recent log file in logs/ directory."""
    if not LOGS_FOLDER.exists():
        return None
    # Today's log first, then most recent
    today_str = datetime.now().strftime("%Y%m%d")
    today_log = LOGS_FOLDER / f"drawingai_{today_str}.log"
    if today_log.exists():
        return today_log
    log_files = sorted(LOGS_FOLDER.glob("drawingai_*.log"), reverse=True)
    return log_files[0] if log_files else None


STATUS_LOG = LOG_DIR / "status_log.txt"

import re as _re
_ANSI_RE = _re.compile(r'\x1b\[[0-9;]*m')


def _parse_log_file_line(line: str) -> str:
    """Convert a drawingai_*.log line to the same format as the console
    (HH:MM:SS │ module │ LEVEL │ message), stripping the date prefix."""
    line = _ANSI_RE.sub('', line)
    parts = line.split(" │ ")
    if len(parts) >= 4:
        timestamp = parts[0].strip()
        # Strip date prefix, keep only HH:MM:SS
        time_short = timestamp.split(" ")[-1] if " " in timestamp else timestamp
        return f"{time_short} │ " + " │ ".join(p.strip() for p in parts[1:])
    return line


def _extract_timestamp(line: str) -> str:
    """Extract HH:MM:SS timestamp from either [HH:MM:SS] or HH:MM:SS │ ... format."""
    line = line.strip()
    if line.startswith("[") and "]" in line:
        return line[1:line.index("]")]
    # Logger format: HH:MM:SS │ ...
    if " │ " in line and len(line) >= 8:
        candidate = line[:8]
        if len(candidate) == 8 and candidate[2] == ':' and candidate[5] == ':':
            return candidate
    return "99:99:99"


def read_log_tail(n_lines: int = 100) -> str:
    """Read status_log.txt which contains both print() output and logger output,
    providing the same detail level as the terminal."""
    lines: list[str] = []

    # ── Primary source: status_log.txt (contains both print + logger output) ──
    if STATUS_LOG.exists():
        try:
            raw = STATUS_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
            for line in raw:
                clean = _ANSI_RE.sub('', line).strip()
                if clean:
                    lines.append(clean)
        except Exception:
            pass

    # ── Fallback: drawingai_*.log (if status_log is empty, e.g. first run) ──
    if not lines:
        log_file = get_latest_log_file()
        if log_file:
            try:
                raw = log_file.read_text(encoding="utf-8", errors="replace").splitlines()
                for line in raw:
                    if line.strip():
                        lines.append(_parse_log_file_line(line))
            except Exception:
                pass

    if not lines:
        return "(אין קבצי לוג)"

    # ── Sort by timestamp ──
    lines.sort(key=_extract_timestamp)

    # ── Deduplicate exact lines ──
    seen: set = set()
    deduped: list[str] = []
    for line in lines:
        if line not in seen:
            seen.add(line)
            deduped.append(line)

    tail = deduped[-n_lines:]
    return "\n".join(tail)


def detect_active_run() -> dict:
    """Detect active run from status_log.txt timestamps and keywords.
    Returns dict with active (bool), run_type ('heavy'|'regular'|''), email_progress ('3/5'|'').
    Works even when session_state.runner is None (e.g. after browser refresh)."""
    import re
    result = {"active": False, "run_type": "", "email_progress": ""}
    if not STATUS_LOG.exists():
        return result
    try:
        mtime = STATUS_LOG.stat().st_mtime
        age = datetime.now().timestamp() - mtime
        if age > 120:  # File not modified in last 2 minutes
            return result

        raw = STATUS_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
        tail = [l.strip() for l in raw[-60:] if l.strip()]
        if not tail:
            return result

        # Check latest timestamp — is it recent?
        now = datetime.now()
        latest_ts = None
        for line in reversed(tail):
            m = re.search(r'(\d{1,2}:\d{2}:\d{2})', line)
            if m:
                try:
                    t = datetime.strptime(m.group(1), "%H:%M:%S").replace(
                        year=now.year, month=now.month, day=now.day)
                    if abs((now - t).total_seconds()) < 120:
                        latest_ts = t
                    break
                except ValueError:
                    pass

        if latest_ts is None:
            return result

        # Recent activity detected — check keywords
        combined = "\n".join(tail[-40:])
        is_heavy = any(k in combined for k in ("כבדים", "HEAVY", "🏋️"))
        is_processing = any(k in combined for k in (
            "מוריד מייל", "מריץ ניתוח", "Stage ",
            "Estimated dimensions", "Extracting high-res",
            "pdfplumber", "Azure Vision",
        ))
        # Check last 3 lines for definitive completion markers
        # Only "הסבב הושלם" or "אין מיילים חדשים בכל התיבות" are real end markers
        # NOTE: "אין מיילים כבדים" per-mailbox is NOT done!
        is_done = False
        for check_line in tail[-3:]:
            if "הסבב הושלם" in check_line:
                is_done = True
                break
            if "אין מיילים חדשים בכל" in check_line:
                is_done = True
                break
            if "הסבב הבא" in check_line:
                is_done = True
                break

        if is_processing and not is_done:
            result["active"] = True
            result["run_type"] = "heavy" if is_heavy else "regular"
            # Try to extract email progress
            for rl in reversed(tail[-30:]):
                pm = re.search(r'מוריד מייל (\d+)/(\d+)', rl)
                if not pm:
                    pm = re.search(r'מריץ ניתוח \[(\d+)/(\d+)\]', rl)
                if pm:
                    result["email_progress"] = f"{pm.group(1)}/{pm.group(2)}"
                    break
    except Exception:
        pass
    return result


def get_countdown() -> dict:
    """Calculate countdown to next automation run from state + config files.
    Returns dict with next_run (str), remaining_seconds (int), remaining_text (str)."""
    import json
    state_path = LOG_DIR / "automation_state.json"
    config_path = LOG_DIR / "automation_config.json"
    result = {"next_run": "", "remaining_seconds": -1, "remaining_text": ""}
    try:
        state = json.loads(state_path.read_text(encoding="utf-8"))
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        return result

    last_checked = state.get("last_checked", "")
    interval = max(int(config.get("poll_interval_minutes", 10)), 1)

    if not last_checked:
        return result

    try:
        # Parse ISO timestamp (may have Z or +00:00)
        lc = last_checked.replace("Z", "+00:00")
        last_dt = datetime.fromisoformat(lc).replace(tzinfo=None)
    except (ValueError, TypeError):
        return result

    from datetime import timedelta
    next_dt = last_dt + timedelta(minutes=interval)
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    remaining = (next_dt - now).total_seconds()

    if remaining < 0:
        result["remaining_text"] = "ממתין לריצה..."
        result["remaining_seconds"] = 0
    else:
        mins = int(remaining) // 60
        secs = int(remaining) % 60
        result["remaining_text"] = f"{mins}:{secs:02d}"
        result["remaining_seconds"] = int(remaining)

    result["next_run"] = next_dt.strftime("%H:%M:%S")
    return result
