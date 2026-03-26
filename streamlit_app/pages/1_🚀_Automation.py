"""
🚀 Automation — DrawingAI Pro
Full automation control panel with tabs: Email, Stages, Run Settings, Live Log
"""
import streamlit as st
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env")

from streamlit_app.backend.config_manager import load_config, save_config, reset_state
from streamlit_app.backend.email_helpers import (
    parse_mailboxes_text, test_all_mailboxes,
    load_folders_for_mailbox, format_folder_label,
)
from streamlit_app.backend.runner_bridge import RunnerBridge, get_runner_bridge
from streamlit_app.backend.log_reader import read_log_tail, get_countdown, detect_active_run
from streamlit_app.brand import BRAND_CSS, brand_header, sidebar_logo


def _is_time_in_range(now_time, start_time, end_time) -> bool:
    """Check if now_time is within [start_time, end_time). Supports overnight ranges (e.g. 19:00→07:00)."""
    if start_time <= end_time:
        return start_time <= now_time < end_time
    else:  # overnight
        return now_time >= start_time or now_time < end_time


st.set_page_config(page_title="🚀 אוטומציה — DrawingAI Pro", page_icon="🌿", layout="wide")
st.markdown(BRAND_CSS, unsafe_allow_html=True)
sidebar_logo()

# ═══════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════
for _k, _d in [
    ("runner", None), ("is_running", False),
    ("folders_loaded", []), ("folders_mailbox", ""),
    ("status_msg", "מוכן"), ("log_lines", []),
    ("confirm_reset", False),
]:
    if _k not in st.session_state:
        st.session_state[_k] = _d

cfg = load_config()

# ═══════════════════════════════════════════════════════════════════
# MODULE-LEVEL CONSTANTS
# ═══════════════════════════════════════════════════════════════════
_CATEGORY_COLOR_MAP = {
    "None": "preset0", "Red": "preset1", "Orange": "preset2",
    "Brown": "preset3", "Yellow": "preset4", "Green": "preset5",
    "Teal": "preset6", "Olive": "preset7", "Blue": "preset8",
    "Purple": "preset9", "Pink": "preset10", "Gray": "preset11",
    "Dark Red": "preset12", "Dark Orange": "preset13", "Dark Brown": "preset14",
    "Dark Yellow": "preset15", "Dark Green": "preset16", "Dark Teal": "preset17",
    "Dark Olive": "preset18", "Dark Blue": "preset19", "Dark Purple": "preset20",
    "Dark Pink": "preset21", "Dark Gray": "preset22", "Black": "preset23",
    "Light Gray": "preset24", "Light Blue": "preset25",
}
_REVERSE_COLOR_MAP = {v: k for k, v in _CATEGORY_COLOR_MAP.items()}
_COLOR_NAMES = list(_CATEGORY_COLOR_MAP.keys())

_RESOLUTION_MAP = {
    "2048 (מהיר - OCR טוב)": 2048,
    "3072 (מאזן - איכות מעולה)": 3072,
    "4096 (איכות - OCR מושלם)": 4096,
    "12000 (Overkill - ברזולוציה מקסימה)": 12000,
}
_REVERSE_RES_MAP = {v: k for k, v in _RESOLUTION_MAP.items()}

_AVAILABLE_MODELS = sorted({
    "gpt-4o-vision", "gpt-4o-mini-email", "o4-mini", "gpt-5.2", "gpt-5.4"
})
_env_path = PROJECT_ROOT / ".env"
if _env_path.exists():
    try:
        for line in _env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            value = value.strip()
            if value and key.startswith("STAGE_") and key.endswith("_MODEL"):
                _AVAILABLE_MODELS.append(value)
            elif key == "AZURE_OPENAI_DEPLOYMENT" and value:
                _AVAILABLE_MODELS.append(value)
        _AVAILABLE_MODELS = sorted(set(_AVAILABLE_MODELS))
    except Exception:
        pass

_ENV_STAGE_DEFAULTS = {}
if _env_path.exists():
    try:
        for line in _env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            value = value.strip()
            if value and key.startswith("STAGE_") and key.endswith("_MODEL"):
                try:
                    stage_num = int(key.split("_")[1])
                    _ENV_STAGE_DEFAULTS[stage_num] = value
                except (ValueError, IndexError):
                    pass
    except Exception:
        pass
for _n in range(10):
    if _n not in _ENV_STAGE_DEFAULTS:
        _ENV_STAGE_DEFAULTS[_n] = "gpt-4o-vision"

# ═══════════════════════════════════════════════════════════════════
# BRANDED HEADER + STATUS
# ═══════════════════════════════════════════════════════════════════
st.html(brand_header("אוטומציה — DrawingAI Pro"))

@st.fragment(run_every=5)
def _header_status_fragment():
    """Auto-refreshing header status — reruns every 5s independently."""
    # Sync is_running with actual thread state
    _r = st.session_state.runner
    if _r and hasattr(_r, 'is_loop_alive'):
        if _r.is_loop_alive and not st.session_state.is_running:
            st.session_state.is_running = True
        elif not _r.is_loop_alive and st.session_state.is_running:
            st.session_state.is_running = False
    _runner = st.session_state.runner
    _log_detect = detect_active_run()

    # ── Determine run states ──
    _regular_busy = _runner and _runner.is_busy and _runner._busy_run_type == "regular"
    _log_regular = _log_detect["active"] and _log_detect["run_type"] == "regular"
    _heavy_busy = _runner and _runner.is_busy and _runner._busy_run_type == "heavy"
    _log_heavy = _log_detect["active"] and _log_detect["run_type"] == "heavy"
    _is_regular_active = st.session_state.is_running or _regular_busy or _log_regular
    _is_heavy_active = _heavy_busy or _log_heavy

    # ── Get email progress ──
    _email_cur, _email_total = 0, 0
    _active_run_type = ""
    if _is_heavy_active:
        _active_run_type = "heavy"
        if _heavy_busy and _runner:
            _s = _runner.get_run_status()
            if _s and _s["total_emails"] > 0:
                _email_cur, _email_total = _s["current_email"], _s["total_emails"]
        elif _log_heavy and _log_detect["email_progress"]:
            _parts = _log_detect["email_progress"].split("/")
            if len(_parts) == 2:
                _email_cur, _email_total = int(_parts[0]), int(_parts[1])
    elif _is_regular_active:
        _active_run_type = "regular"
        if _regular_busy and _runner:
            _s = _runner.get_run_status()
            if _s and _s["total_emails"] > 0:
                _email_cur, _email_total = _s["current_email"], _s["total_emails"]
        elif _log_regular and _log_detect["email_progress"]:
            _parts = _log_detect["email_progress"].split("/")
            if len(_parts) == 2:
                _email_cur, _email_total = int(_parts[0]), int(_parts[1])

    # ── Build progress pill ──
    _progress_html = ""
    if _email_total > 0 and (_is_regular_active or _is_heavy_active):
        _progress_html = f'<span class="email-progress">📧 {_email_cur}/{_email_total}</span>'

    _cols = st.columns([3, 1, 1])
    with _cols[0]:
        # ── Regular automation badge ──
        if st.session_state.is_running:
            _hdr_cd = get_countdown()
            _cd_text = f' — ⏱ סבב הבא: {_hdr_cd["remaining_text"]}' if _hdr_cd["remaining_text"] else ""
            _reg_progress = _progress_html if _active_run_type == "regular" else ""
            st.markdown(f'<span class="run-badge run-badge-active">🟢 אוטומציה רגילה: פעילה{_cd_text}</span> {_reg_progress}', unsafe_allow_html=True)
        elif _regular_busy or _log_regular:
            st.markdown(f'<span class="run-badge run-badge-active">🟢 ריצה רגילה: פעילה</span> {_progress_html if _active_run_type == "regular" else ""}', unsafe_allow_html=True)
        else:
            st.markdown('<span class="run-badge run-badge-inactive">🔴 אוטומציה רגילה: כבויה</span>', unsafe_allow_html=True)

        # ── Heavy run badge ──
        if _is_heavy_active:
            st.markdown(f'<span class="run-badge run-badge-heavy">🏋️ ריצה כבדה: פעילה</span> {_progress_html if _active_run_type == "heavy" else ""}', unsafe_allow_html=True)
        else:
            st.markdown('<span class="run-badge run-badge-inactive">⚪ ריצה כבדה: לא פעילה</span>', unsafe_allow_html=True)

        # ── Scheduler status ──
        _sched_runner = st.session_state.runner
        if _sched_runner and _sched_runner.scheduler_active:
            st.markdown('<span class="status-running">🕐 תזמון: פעיל</span>', unsafe_allow_html=True)
        elif cfg.get("scheduler_enabled", False):
            st.markdown('<span class="status-stopped" style="opacity:0.5;">🕐 תזמון: מוגדר (לא פעיל)</span>', unsafe_allow_html=True)

    # ── Cost display ──
    with _cols[1]:
        _cost_text = "—"
        try:
            from streamlit_app.backend.log_reader import load_log_entries, filter_by_period
            _today_entries = filter_by_period(load_log_entries(500), "היום")
            _today_cost = sum(float(e.get("cost_usd", 0) or 0) for e in _today_entries)
            if _today_cost > 0:
                _ils_rate = float(cfg.get("usd_to_ils_rate", 3.7))
                _cost_text = f'${_today_cost:.3f} (₪{_today_cost * _ils_rate:.2f})'
        except Exception:
            pass
        st.markdown(f'<div style="text-align:center;font-size:13px;">💰 היום: {_cost_text}</div>', unsafe_allow_html=True)

_header_status_fragment()

# ═══════════════════════════════════════════════════════════════════
# ACTION BUTTONS (always visible above tabs)
# ═══════════════════════════════════════════════════════════════════
btn_cols = st.columns(7)
with btn_cols[0]:
    save_btn = st.button("💾 שמור", width='stretch')
with btn_cols[1]:
    test_btn = st.button("🔌 בדוק חיבור", width='stretch')
with btn_cols[2]:
    run_once_btn = st.button("▶️ הרץ סבב", width='stretch')
with btn_cols[3]:
    run_heavy_btn = st.button("🏋️ הרץ כבדים", width='stretch')
with btn_cols[4]:
    start_btn = st.button("🚀 הפעל", width='stretch')
with btn_cols[5]:
    stop_btn = st.button("⏹ עצור", width='stretch')
with btn_cols[6]:
    if st.session_state.confirm_reset:
        reset_btn = st.button("⚠️ לחץ שוב לאישור", width='stretch', type="primary")
    else:
        reset_btn = st.button("🔄 Reset", width='stretch')

# Local save button for scheduler section (set later inside tab_run)
save_scheduler_btn = False

# ═══════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════
tab_email, tab_stages, tab_run, tab_log = st.tabs([
    "📧 מייל ותיקיות", "🔬 שלבים ומודלים", "⏱ הגדרות ריצה", "📋 לוג ריצה"
])

# ╔═══════════════════════════════════════════════════════════╗
# ║  TAB 1 — EMAIL + FOLDERS                                 ║
# ╚═══════════════════════════════════════════════════════════╝
with tab_email:
    col_mail, col_folders = st.columns([1, 1])

    with col_mail:
        st.subheader("📧 הגדרות מייל")

        shared_mailboxes_str = st.text_input(
            "תיבות משותפות (מופרדות בפסיק):",
            value=", ".join(cfg.get("shared_mailboxes", [])) or cfg.get("shared_mailbox", ""),
            help="כתובות תיבות משותפות, מופרדות בפסיקים. ניתן להזין מספר תיבות.",
            key="shared_mailboxes_input",
        )
        parsed_mailboxes = parse_mailboxes_text(shared_mailboxes_str)

        mail_c1, mail_c2 = st.columns([3, 1])
        with mail_c1:
            selected_mailbox = st.selectbox(
                "תיבה להצגה:", options=parsed_mailboxes if parsed_mailboxes else [""],
                index=0, key="selected_mailbox",
            )
        with mail_c2:
            st.markdown("<br>", unsafe_allow_html=True)
            load_folders_btn = st.button("📂 טען תיקיות", width='stretch', key="load_folders_btn")

        if load_folders_btn and selected_mailbox:
            with st.spinner(f"טוען תיקיות עבור {selected_mailbox}..."):
                folders, err = load_folders_for_mailbox(selected_mailbox)
                if err:
                    st.error(err)
                else:
                    st.session_state.folders_loaded = folders
                    st.session_state.folders_mailbox = selected_mailbox
                    st.success(f"נטענו {len(folders)} תיקיות עבור {selected_mailbox}")

        folder_options = []
        folder_path_map = {}
        if st.session_state.folders_loaded:
            for f in st.session_state.folders_loaded:
                label = format_folder_label(f["path"], f.get("totalItemCount"))
                folder_options.append(label)
                folder_path_map[label] = f["path"]

        current_folder = cfg.get("folder_name", "Inbox")
        default_index = 0
        if folder_options:
            for i, opt in enumerate(folder_options):
                raw_path = folder_path_map.get(opt, opt)
                if raw_path == current_folder or opt.startswith(current_folder):
                    default_index = i
                    break
            selected_folder_label = st.selectbox(
                "תת-תיקייה:", options=folder_options, index=default_index, key="folder_select",
            )
            selected_folder = folder_path_map.get(selected_folder_label, selected_folder_label)
        else:
            selected_folder = st.text_input("תת-תיקייה:", value=current_folder, key="folder_text")

        rerun_c1, rerun_c2 = st.columns(2)
        with rerun_c1:
            rerun_folder = st.text_input("תיקיית RERUN:", value=cfg.get("rerun_folder_name", ""), key="rerun_folder",
                                          help="שם תיקייה ב-Outlook להרצה חוזרת. גררי מייל לכאן כדי לשלוח שוב עם כל הפריטים.")
        with rerun_c2:
            rerun_mailbox = st.text_input("תיבת RERUN:", value=cfg.get("rerun_mailbox", ""), key="rerun_mailbox",
                                           help="תיבה בה נמצאת תיקיית RERUN. השאירי ריק אם היא באותה תיבה ראשית.")

        scan_from_date = st.text_input(
            "סרוק מתאריך (DD/MM/YYYY HH:MM):", value=cfg.get("scan_from_date", ""), key="scan_from_date",
            help="מיילים שהתקבלו לפני תאריך זה יתעלמו. השאירי ריק לסריקה רגילה.",
        )
        recipient_email = st.text_input("נמען לשליחה:", value=cfg.get("recipient_email", ""), key="recipient_email",
                                         help="כתובת המייל שאליה יישלחו קבצי B2B.")

        cat_c1, cat_c2 = st.columns(2)
        with cat_c1:
            mark_cat_name = st.text_input("קטגוריה לסימון:", value=cfg.get("mark_category_name", "AI Processed"), key="mark_cat_name")
        with cat_c2:
            stored_color = cfg.get("mark_category_color", "preset20")
            mark_cat_color = st.selectbox(
                "צבע:", options=_COLOR_NAMES,
                index=_COLOR_NAMES.index(_REVERSE_COLOR_MAP.get(stored_color, "None")), key="mark_cat_color",
            )
        cat_c3, cat_c4 = st.columns(2)
        with cat_c3:
            nodraw_cat_name = st.text_input("קטגוריה NO DRAW:", value=cfg.get("nodraw_category_name", "NO DRAW"), key="nodraw_cat_name")
        with cat_c4:
            stored_nodraw_color = cfg.get("nodraw_category_color", "preset1")
            nodraw_cat_color = st.selectbox(
                "צבע NO DRAW:", options=_COLOR_NAMES,
                index=_COLOR_NAMES.index(_REVERSE_COLOR_MAP.get(stored_nodraw_color, "None")), key="nodraw_cat_color",
            )

        skip_senders = st.text_area(
            "דלג על שולחים (כתובת בכל שורה):",
            value="\n".join(cfg.get("skip_senders", [])), height=80, key="skip_senders",
            help="כתובות מייל שמהן לא לעבד. כתובת אחת בכל שורה.",
        )
        skip_categories = st.text_area(
            "דלג על קטגוריות (שם בכל שורה):",
            value="\n".join(cfg.get("skip_categories", [])), height=60, key="skip_categories",
            help="מיילים עם קטגוריות אלו ב-Outlook ידולגו.",
        )

    with col_folders:
        st.subheader("📁 תיקיות")
        download_root = st.text_input("תיקיית הורדה:", value=cfg.get("download_root", ""), key="download_root",
                                        help="תיקייה להורדת קבצים מצורפים ממיילים.")
        if download_root and not Path(download_root).exists():
            st.warning("⚠️ התיקייה לא קיימת")
        tosend_folder = st.text_input("TO_SEND:", value=cfg.get("tosend_folder", ""), key="tosend_folder",
                                        help="תיקייה שבה יישמרו קבצי B2B המוכנים לשליחה.")
        if tosend_folder and not Path(tosend_folder).exists():
            st.warning("⚠️ התיקייה לא קיימת")
        output_copy_folder = st.text_input("תיקיית שמירה:", value=cfg.get("output_copy_folder", ""), key="output_copy_folder",
                                             help="תיקייה לשמירת עותק מלא של הקבצים המעובדים.")

# ╔═══════════════════════════════════════════════════════════╗
# ║  TAB 2 — STAGES + MODELS                                 ║
# ╚═══════════════════════════════════════════════════════════╝
with tab_stages:
    st.subheader("🔬 שלבי חילוץ ומודלים")

    stage_defs = [
        (0, "0: זיהוי", False), (1, "1: בסיסי", True),
        (2, "2: תהליכים", True), (3, "3: NOTES", True),
        (4, "4: שטח", True), (5, "5: Fallback", True),
        (6, "6: PL", True), (7, "7: email", True),
        (8, "8: הזמנות", True), (9, "9: מיזוג", True),
    ]
    saved_stages = cfg.get("selected_stages", {})
    saved_models = cfg.get("stage_models", {})
    stage_enabled = {}
    stage_models = {}

    cols_r1 = st.columns(5)
    for i, (sn, label, has_cb) in enumerate(stage_defs[:5]):
        with cols_r1[i]:
            if has_cb:
                stage_enabled[sn] = st.checkbox(label, value=bool(saved_stages.get(str(sn), True)), key=f"stage_{sn}_on")
            else:
                st.markdown(f"**{label}**")
                stage_enabled[sn] = True
            md = saved_models.get(str(sn), _ENV_STAGE_DEFAULTS.get(sn, "gpt-4o-vision"))
            mi = _AVAILABLE_MODELS.index(md) if md in _AVAILABLE_MODELS else 0
            stage_models[sn] = st.selectbox("מודל", options=_AVAILABLE_MODELS, index=mi, key=f"model_{sn}", label_visibility="collapsed")

    cols_r2 = st.columns(5)
    for i, (sn, label, has_cb) in enumerate(stage_defs[5:]):
        with cols_r2[i]:
            if has_cb:
                stage_enabled[sn] = st.checkbox(label, value=bool(saved_stages.get(str(sn), True)), key=f"stage_{sn}_on")
            md = saved_models.get(str(sn), _ENV_STAGE_DEFAULTS.get(sn, "gpt-4o-vision"))
            mi = _AVAILABLE_MODELS.index(md) if md in _AVAILABLE_MODELS else 0
            stage_models[sn] = st.selectbox("מודל", options=_AVAILABLE_MODELS, index=mi, key=f"model_{sn}", label_visibility="collapsed")

# ╔═══════════════════════════════════════════════════════════╗
# ║  TAB 3 — RUN SETTINGS                                    ║
# ╚═══════════════════════════════════════════════════════════╝
with tab_run:
    st.subheader("⏱ הגדרות ריצה")

    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        poll_interval = st.number_input("דקות בין סבבים:", value=int(cfg.get("poll_interval_minutes", 10)), min_value=1, max_value=120, key="poll_interval",
                                        help="כל כמה דקות לבדוק מיילים חדשים.")
    with rc2:
        max_messages = st.number_input("כמות מיילים לסבב:", value=int(cfg.get("max_messages", 200)), min_value=1, max_value=5000, key="max_messages",
                                        help="כמה מיילים לעבד בסבב אחד (מקסימום).")
    with rc3:
        max_files = st.number_input("מקסימום קבצים למייל (0=ללא):", value=int(cfg.get("max_files_per_email", 15)), min_value=0, max_value=100, key="max_files",
                                     help="מיילים עם יותר קבצים מהסף יסומנו AI HEAVY וידלגו. 0=ללא הגבלה.")

    chk = st.columns(5)
    with chk[0]:
        auto_start = st.checkbox("הפעל בעת פתיחה", value=cfg.get("auto_start", False), key="auto_start",
                                  help="התחל ריצה אוטומטית כשפותחים את האפליקציה.")
    with chk[1]:
        auto_send = st.checkbox("שלח אוטומטית", value=cfg.get("auto_send", False), key="auto_send",
                                 help="שלח את קבצי B2B במייל אוטומטית לאחר עיבוד.")
    with chk[2]:
        archive_full = st.checkbox("שמור עותק מלא", value=cfg.get("archive_full", False), key="archive_full",
                                    help="שמור עותק מלא של כל הקבצים בתיקיית שמירה.")
    with chk[3]:
        cleanup_download = st.checkbox("מחק אחרי העברה", value=cfg.get("cleanup_download", True), key="cleanup_download",
                                        help="נקה קבצים שהורדו לאחר העברתם לתיקיית הפלט.")
    with chk[4]:
        mark_processed = st.checkbox("סמן מעובד", value=cfg.get("mark_as_processed", True), key="mark_processed",
                                      help="סמן מיילים כ'מעובד' ב-Outlook כדי לא לעבד שוב.")

    with st.expander("🕐 תזמון אוטומטי (Scheduler)", expanded=cfg.get("scheduler_enabled", False)):
        st.caption("הגדר שעות ריצה אוטומטית — רגילה וכבדה ישלבו לפי טווח השעות. כשהתזמון פועל, המערכת מריצה סבבים לפי הזמן ללא צורך בלחיצה ידנית.")

        sched_enabled = st.checkbox("הפעל תזמון אוטומטי", value=cfg.get("scheduler_enabled", False), key="sched_enabled",
                                     help="כשמופעל, המערכת תריץ סבבים אוטומטית בטווח השעות שנקבע.")

        sc1, sc2 = st.columns(2)
        with sc1:
            st.markdown("**▶️ רגילה**")
            from datetime import time as dt_time
            _reg_from = cfg.get("scheduler_regular_from", "07:00")
            _reg_to = cfg.get("scheduler_regular_to", "19:00")
            sched_regular_from = st.time_input("משעה:", value=dt_time(*map(int, _reg_from.split(":"))), key="sched_reg_from")
            sched_regular_to = st.time_input("עד שעה:", value=dt_time(*map(int, _reg_to.split(":"))), key="sched_reg_to")
        with sc2:
            st.markdown("**🏋️ כבדה**")
            _hvy_from = cfg.get("scheduler_heavy_from", "19:00")
            _hvy_to = cfg.get("scheduler_heavy_to", "07:00")
            sched_heavy_from = st.time_input("משעה:", value=dt_time(*map(int, _hvy_from.split(":"))), key="sched_hvy_from")
            sched_heavy_to = st.time_input("עד שעה:", value=dt_time(*map(int, _hvy_to.split(":"))), key="sched_hvy_to")

        sched_interval = st.number_input("דקות בין סבבים מתוזמנים:", value=int(cfg.get("scheduler_interval_minutes", 10)),
                                          min_value=1, max_value=120, key="sched_interval",
                                          help="כל כמה דקות לבדוק ולהריץ בטווח המתוזמן.")

        sched_report_folder = st.text_input(
            "📊 תיקיית דוחות Excel (אופציונלי):",
            value=cfg.get("scheduler_report_folder", ""),
            key="sched_report_folder",
            placeholder=r"לדוגמה: C:\Reports\scheduler",
            help="אם מוגדר, יישמר דוח Excel סיכומי בתיקייה זו בתום כל סבב מתוזמן. ריק = ללא דוח.",
        )

        save_scheduler_btn = st.button(
            "💾 שמור הגדרות תזמון",
            key="save_scheduler_btn",
            width='stretch',
            help="שומר מיד את שעות התזמון ונתיב דוח ה-Excel",
        )

        if sched_enabled:
            from datetime import datetime as _dt
            _now = _dt.now().time()
            _in_regular = _is_time_in_range(_now, sched_regular_from, sched_regular_to)
            _in_heavy = _is_time_in_range(_now, sched_heavy_from, sched_heavy_to)
            if _in_regular:
                st.success(f"⏰ כרגע בטווח ריצה **רגילה** (שעה: {_now.strftime('%H:%M')})")
            elif _in_heavy:
                st.success(f"⏰ כרגע בטווח ריצה **כבדה** (שעה: {_now.strftime('%H:%M')})")
            else:
                st.warning(f"⏰ כרגע לא בטווח ריצה (שעה: {_now.strftime('%H:%M')})")

    with st.expander("⚙️ הגדרות מתקדמות"):
        ac1, ac2 = st.columns(2)
        with ac1:
            stored_dim = cfg.get("max_image_dimension", 4096)
            dim_label = _REVERSE_RES_MAP.get(stored_dim, "4096 (איכות - OCR מושלם)")
            max_image_dim = st.selectbox(
                "מקסימום רזולוציה:", options=list(_RESOLUTION_MAP.keys()),
                index=list(_RESOLUTION_MAP.keys()).index(dim_label) if dim_label in _RESOLUTION_MAP else 2,
                key="max_image_dim",
                help="רזולוציה גבוהה יותר = דיוק טוב יותר אך עלות גבוהה יותר.",
            )
            max_file_size = st.number_input("מקסימום גודל קובץ (MB):", value=int(cfg.get("max_file_size_mb", 100)), min_value=1, max_value=500, key="max_file_size",
                                             help="קבצים גדולים מהסף ידלגו.")
            stage1_skip = st.number_input("דילוג Retry שלב 1 (px):", value=int(cfg.get("stage1_skip_retry_resolution_px", 8000)), min_value=0, key="s1_skip",
                                           help="תמונות עם רזולוציה גבוהה מהסף לא עוברות ניסיון חוזר בשלב 1.")
            confidence = st.radio(
                "רמת ביטחון B2B:", options=["LOW", "MEDIUM", "HIGH"],
                index=["LOW", "MEDIUM", "HIGH"].index(cfg.get("confidence_level", "LOW")),
                horizontal=True, key="confidence",
                help="LOW=מקבל גם מסמכים לא ברורים, HIGH=רק בטוחים.",
            )
            recursive = st.checkbox("כולל תת-תיקיות", value=cfg.get("recursive", True), key="recursive",
                                     help="סרוק גם תיקיות פנימיות.")
            enable_retry = st.checkbox("ניסיונות נוספים", value=cfg.get("enable_retry", True), key="enable_retry",
                                        help="אפשר ניסיונות חוזרים כשהעיבוד נכשל.")
        with ac2:
            debug_mode = st.checkbox("Debug mode", value=cfg.get("debug_mode", False), key="debug_mode",
                                      help="הדפסות מפורטות בלוג לצורך ניפוי בעיות.")
            iai_top_red = st.checkbox("IAI top-red fallback", value=cfg.get("iai_top_red_fallback", True), key="iai_top_red",
                                      help="נסה חלופת top-red כאשר זיהוי רגיל נכשל.")
            max_retries = st.number_input("ניסיונות חוזרים:", value=int(cfg.get("max_retries", 3)), min_value=1, max_value=5, key="max_retries",
                                           help="מספר ניסיונות חוזרים לפני דילוג על קובץ.")
            scan_dpi = st.selectbox("DPI:", options=[150, 200, 300], index=[150, 200, 300].index(int(cfg.get("scan_dpi", 200))), key="scan_dpi",
                                     help="רזולוציית סריקה. 300=מומלץ ל-OCR.")
            log_max = st.number_input("לוג מקס (MB):", value=int(cfg.get("log_max_size_mb", 1)), min_value=1, max_value=50, key="log_max",
                                       help="גודל מקסימלי של קובץ לוג לפני רוטציה.")
            usd_to_ils = st.number_input("$/₪:", value=float(cfg.get("usd_to_ils_rate", 3.7)), min_value=0.1, max_value=10.0, step=0.1, key="usd_ils",
                                          help="שער המרה לתצוגת עלות בשקלים.")

# ╔═══════════════════════════════════════════════════════════╗
# ║  TAB 4 — LIVE LOG                                        ║
# ╚═══════════════════════════════════════════════════════════╝
with tab_log:
    st.subheader("📋 לוג ריצה")

    @st.fragment(run_every=5)
    def _live_log_fragment():
        """Auto-refreshing log fragment — reruns every 5s independently."""
        _is_active = st.session_state.is_running or (st.session_state.runner and st.session_state.runner.is_busy)
        _countdown = get_countdown()
        _runner_status = None
        if st.session_state.runner:
            _runner_status = st.session_state.runner.get_run_status()

        _phase_map = {
            "idle": "⏸ ממתין",
            "scanning": "🔍 סורק מיילים...",
            "processing": "⚙️ מעבד...",
            "sending": "📤 שולח מייל...",
            "done": "✅ הסבב הושלם",
            "waiting": "⏳ ממתין לסבב הבא",
        }
        _run_type_labels = {
            "regular": "🔄 ריצה רגילה",
            "heavy": "🏋️ ריצת כבדים",
        }
        _phase_text = ""
        _progress_text = ""
        _timer_text = ""
        _run_type_text = ""

        if _runner_status:
            _phase = _runner_status["phase"]
            _run_type_text = _run_type_labels.get(_runner_status.get("run_type", ""), "")

            if _phase in ("processing", "scanning", "sending"):
                _phase_text = _phase_map.get(_phase, "")
                if _runner_status["total_emails"] > 0:
                    _progress_text = f"📧 מייל {_runner_status['current_email']}/{_runner_status['total_emails']}"
            elif _phase == "done":
                _phase_text = _phase_map["done"]
            elif _runner_status["next_run"]:
                _phase_text = _phase_map["waiting"]

            if _runner_status["next_run"]:
                _timer_text = f"⏱ סבב הבא: {_runner_status['next_run']}"

        # Only show countdown timer if NOT actively processing
        _active_log_phases = ("processing", "scanning", "sending")
        _is_log_active = _runner_status and _runner_status["phase"] in _active_log_phases
        if not _timer_text and _countdown["remaining_text"] and not _is_log_active and not _is_active:
            _timer_text = f"⏱ סבב הבא: {_countdown['remaining_text']}"
            if not _phase_text:
                _phase_text = _phase_map["waiting"]

        # Show "running" in status bar when is_busy even if log hasn't detected it yet
        if _is_active and not _phase_text:
            _phase_text = _phase_map["scanning"]
            if st.session_state.runner and st.session_state.runner._busy_run_type == "heavy":
                _run_type_text = _run_type_labels["heavy"]
            elif st.session_state.runner and st.session_state.runner._busy_run_type == "regular":
                _run_type_text = _run_type_labels["regular"]

        # Always show status bar
        if _phase_text or _timer_text or _progress_text or _run_type_text:
            _bar_parts = []
            if _phase_text:
                _bar_parts.append(f'<span class="phase">{_phase_text}</span>')
            if _run_type_text:
                _bar_parts.append(f'<span class="progress">{_run_type_text}</span>')
            if _progress_text:
                _bar_parts.append(f'<span class="progress">{_progress_text}</span>')
            if _timer_text:
                _bar_parts.append(f'<span class="timer">{_timer_text}</span>')
            st.markdown(
                f'<div class="status-bar">{"".join(_bar_parts)}</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="status-bar"><span class="phase">⏸ לא פעיל</span></div>',
                unsafe_allow_html=True,
            )

        # Merged log: status_log.txt + drawingai_*.log for full Tkinter-level detail
        log_text = read_log_tail(2000)

        st.html(
            f'<div class="log-container" id="logBox" style="'
            f'background:#1a1a2e;color:#00ff00;font-family:Consolas,monospace;'
            f'font-size:12px;padding:12px;border-radius:8px;height:550px;'
            f'overflow-y:scroll;direction:ltr;text-align:left;white-space:pre-wrap;'
            f'border:1px solid #FF8C0033;">{log_text}</div>'
            '<script>var lb=document.getElementById("logBox");if(lb)lb.scrollTop=lb.scrollHeight;</script>'
        )

        btn_cols = st.columns([1, 1, 6])
        with btn_cols[0]:
            if st.button("🗑 נקה לוג", key="clear_log"):
                _status_log = PROJECT_ROOT / "status_log.txt"
                try:
                    _status_log.write_text("", encoding="utf-8")
                except Exception:
                    pass
                if st.session_state.runner:
                    st.session_state.runner.clear_log()
                st.rerun()
        with btn_cols[1]:
            if st.button("🔄 רענן", key="refresh_log"):
                st.rerun()

    _live_log_fragment()

# ═══════════════════════════════════════════════════════════════════
# BUTTON HANDLERS (outside tabs)
# ═══════════════════════════════════════════════════════════════════

def _gather_config() -> dict:
    pm = parse_mailboxes_text(shared_mailboxes_str)
    primary = pm[0] if pm else ""
    fv = selected_folder
    if fv and "(" in fv:
        fv = folder_path_map.get(fv, fv)
    return {
        "shared_mailbox": primary,
        "shared_mailboxes": pm,
        "folder_name": fv.strip() if fv else "Inbox",
        "rerun_folder_name": rerun_folder.strip(),
        "rerun_mailbox": rerun_mailbox.strip(),
        "skip_senders": [s.strip().lower() for s in skip_senders.strip().splitlines() if s.strip()],
        "skip_categories": [c.strip() for c in skip_categories.strip().splitlines() if c.strip()],
        "scan_from_date": scan_from_date.strip(),
        "recipient_email": recipient_email.strip(),
        "download_root": download_root.strip(),
        "tosend_folder": tosend_folder.strip(),
        "output_copy_folder": output_copy_folder.strip(),
        "poll_interval_minutes": poll_interval,
        "max_messages": max_messages,
        "max_files_per_email": max_files,
        "max_file_size_mb": max_file_size,
        "stage1_skip_retry_resolution_px": stage1_skip,
        "max_image_dimension": _RESOLUTION_MAP.get(max_image_dim, 4096),
        "recursive": recursive,
        "enable_retry": enable_retry,
        "auto_start": auto_start,
        "auto_send": auto_send,
        "archive_full": archive_full,
        "cleanup_download": cleanup_download,
        "mark_as_processed": mark_processed,
        "mark_category_name": mark_cat_name.strip() or "AI Processed",
        "mark_category_color": _CATEGORY_COLOR_MAP.get(mark_cat_color, "preset0"),
        "nodraw_category_name": nodraw_cat_name.strip() or "NO DRAW",
        "nodraw_category_color": _CATEGORY_COLOR_MAP.get(nodraw_cat_color, "preset1"),
        "heavy_category_name": "AI HEAVY",
        "heavy_category_color": "preset4",
        "confidence_level": confidence,
        "debug_mode": debug_mode,
        "iai_top_red_fallback": iai_top_red,
        "max_retries": max_retries,
        "scan_dpi": scan_dpi,
        "log_max_size_mb": log_max,
        "usd_to_ils_rate": usd_to_ils,
        "selected_stages": {str(n): stage_enabled.get(n, True) for n in range(1, 10)},
        "stage_models": {str(n): stage_models.get(n, "") for n in range(10) if stage_models.get(n)},
        "scheduler_enabled": sched_enabled,
        "scheduler_regular_from": sched_regular_from.strftime("%H:%M"),
        "scheduler_regular_to": sched_regular_to.strftime("%H:%M"),
        "scheduler_heavy_from": sched_heavy_from.strftime("%H:%M"),
        "scheduler_heavy_to": sched_heavy_to.strftime("%H:%M"),
        "scheduler_interval_minutes": sched_interval,
        "scheduler_report_folder": sched_report_folder.strip(),
    }


if save_btn or save_scheduler_btn:
    new_cfg = _gather_config()
    save_config(new_cfg)
    st.toast("💾 ההגדרות נשמרו בהצלחה!", icon="✅")

if test_btn:
    if not parsed_mailboxes:
        st.error("לא הוגדרה אף תיבה לבדיקה")
    else:
        with st.spinner("בודק חיבור לתיבות..."):
            results = test_all_mailboxes(parsed_mailboxes)
            ok = [m for m, s in results.items() if s]
            fail = [m for m, s in results.items() if not s]
        if ok:
            st.success(f"✅ תיבות תקינות ({len(ok)}): " + ", ".join(ok))
        if fail:
            st.error(f"❌ תיבות שנכשלו ({len(fail)}): " + ", ".join(fail))

if run_once_btn:
    new_cfg = _gather_config()
    save_config(new_cfg)
    if st.session_state.runner is None:
        st.session_state.runner = get_runner_bridge()
    _runner = st.session_state.runner
    if _runner.is_busy or _runner.is_running:
        st.toast("⚠️ ריצה כבר פעילה — לא ניתן להריץ במקביל", icon="⚠️")
    elif _runner.run_once():
        st.toast("▶️ מריץ סבב אחד...", icon="▶️")
    else:
        st.toast("⚠️ ריצה כבר פעילה — לא ניתן להריץ במקביל", icon="⚠️")
    st.rerun()

if run_heavy_btn:
    new_cfg = _gather_config()
    save_config(new_cfg)
    if st.session_state.runner is None:
        st.session_state.runner = get_runner_bridge()
    _runner = st.session_state.runner
    if _runner.is_busy or _runner.is_running:
        st.toast("⚠️ ריצה כבר פעילה — לא ניתן להריץ במקביל", icon="⚠️")
    elif _runner.run_heavy():
        st.toast("🏋️ מריץ כבדים...", icon="🏋️")
    else:
        st.toast("⚠️ ריצה כבר פעילה — לא ניתן להריץ במקביל", icon="⚠️")
    st.rerun()

# Show progress indicator when a one-shot run is active
_r = st.session_state.runner
if _r and _r.is_busy:
    _type_label = "סבב רגיל" if _r._busy_run_type == "regular" else "ריצת כבדים"
    with st.status(f"⏳ מריץ {_type_label}...", expanded=False):
        st.write("הריצה פעילה. עקוב אחרי הלוג לפרטים.")

if start_btn:
    new_cfg = _gather_config()
    save_config(new_cfg)
    if st.session_state.runner is None:
        st.session_state.runner = get_runner_bridge()
    _runner = st.session_state.runner
    if _runner.is_busy:
        st.toast("⚠️ ריצה חד-פעמית עדיין פעילה — לא ניתן להפעיל לולאה", icon="⚠️")
    elif _runner.start():
        st.session_state.is_running = True
        st.toast("🚀 אוטומציה הופעלה!", icon="🚀")
    st.rerun()

if stop_btn:
    if st.session_state.runner:
        st.session_state.runner.stop()
        st.session_state.runner.stop_scheduler()
    st.session_state.is_running = False
    st.toast("⏹ אוטומציה נעצרה", icon="⏹")
    st.rerun()

# ── Auto-start/stop scheduler based on checkbox ──
if sched_enabled:
    if st.session_state.runner is None:
        st.session_state.runner = get_runner_bridge()
    _sched_runner = st.session_state.runner
    if not _sched_runner.scheduler_active:
        new_cfg = _gather_config()
        save_config(new_cfg)
        _sched_runner.start_scheduler()
else:
    _sched_runner = st.session_state.runner
    if _sched_runner and _sched_runner.scheduler_active:
        _sched_runner.stop_scheduler()

if reset_btn:
    if st.session_state.confirm_reset:
        reset_state()
        st.session_state.confirm_reset = False
        st.toast("🔄 State אופס — כל המיילים יתחשבו כחדשים", icon="🔄")
        st.rerun()
    else:
        st.session_state.confirm_reset = True
        st.warning("⚠️ פעולה זו תאפס את המצב — כל המיילים יעובדו מחדש. לחץ שוב לאישור.")
        st.rerun()
