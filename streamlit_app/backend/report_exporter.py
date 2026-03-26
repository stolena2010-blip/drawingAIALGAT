"""Scheduled report export wrapper using the shared Excel builder."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from streamlit_app.backend.excel_report_builder import build_workbook_scheduler
from streamlit_app.backend.log_reader import load_log_entries


def _entry_dt(e: Dict[str, Any]) -> Optional[datetime]:
    ts = e.get("timestamp") or e.get("start_time") or ""
    try:
        return datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone().replace(tzinfo=None)
    except Exception:
        return None


def export_schedule_report(output_folder: str, run_type: str, run_start=None) -> str:
    """Generate scheduler report and keep only one latest file in output_folder."""
    folder = Path(output_folder)
    folder.mkdir(parents=True, exist_ok=True)

    run_label = "regular" if run_type == "regular" else "heavy"
    filepath = folder / "schedule_report_latest.xlsx"

    # Remove historical scheduler reports so only one file remains.
    for old_file in folder.glob("schedule_report_*.xlsx"):
        try:
            old_file.unlink()
        except Exception:
            pass

    all_entries = load_log_entries()
    if run_start is not None:
        run_entries = [
            e for e in all_entries
            if (edt := _entry_dt(e)) is not None and edt >= run_start
        ]
    else:
        run_entries = all_entries

    wb = build_workbook_scheduler(
        all_entries=all_entries,
        run_entries=run_entries,
        run_label=run_label,
        run_start=run_start,
    )
    wb.save(str(filepath))
    return str(filepath)
