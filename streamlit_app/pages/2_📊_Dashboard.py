"""
📊 Dashboard — DrawingAI Pro
Full statistics: summary, accuracy, efficiency, customers, senders, recent emails, export
Mirrors original dashboard_gui.py features with tabs for convenience.
"""
import streamlit as st
import sys
import os
import json
import numpy  # force full init before plotly (avoids circular import race)
from pathlib import Path
from datetime import datetime, timedelta, timezone
from collections import Counter, defaultdict

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env")

from streamlit_app.backend.log_reader import (
    load_log_entries, filter_by_period,
    get_accuracy_weights, calc_weighted_accuracy, save_entry_field, LOG_DIR,
)
from streamlit_app.backend.excel_report_builder import build_workbook_dashboard, workbook_to_bytes
from streamlit_app.brand import BRAND_CSS, brand_header, sidebar_logo

st.set_page_config(page_title="📊 Dashboard — DrawingAI Pro", page_icon="🌿", layout="wide")

st.markdown(BRAND_CSS, unsafe_allow_html=True)
sidebar_logo()
# Extra page-specific styles handled by brand.py BRAND_CSS

st.html(brand_header("דשבורד — DrawingAI Pro"))

# ═══════  PERIOD FILTER BAR  ═══════
filter_cols = st.columns([2, 1, 1, 1])
with filter_cols[0]:
    period = st.radio("תקופה:", ["היום", "שבוע", "חודש", "הכל", "טווח..."], horizontal=True, key="pf")
with filter_cols[1]:
    date_from = ""
    if period == "טווח...":
        date_from = st.date_input("מתאריך", value=datetime.now() - timedelta(days=30), key="df").strftime("%Y-%m-%d")
with filter_cols[2]:
    date_to = ""
    if period == "טווח...":
        date_to = st.date_input("עד תאריך", value=datetime.now(), key="dt").strftime("%Y-%m-%d")

# ═══════  LOAD DATA  ═══════
all_entries = load_log_entries()
period_key = period if period != "טווח..." else "טווח"
filtered = filter_by_period(all_entries, period_key, date_from, date_to)

# Skip non-email entries (events, RERUN-only)
email_entries = [e for e in filtered if e.get("accuracy_data") or e.get("files_processed")]

weights = get_accuracy_weights()


# ═══════  HELPER FUNCTIONS  ═══════
def _entry_ts(e):
    ts = e.get("timestamp") or e.get("received") or ""
    try:
        return datetime.fromisoformat(ts.replace("Z", "+00:00")).replace(tzinfo=None)
    except (ValueError, TypeError):
        return datetime.min


def _acc_color(val):
    if val >= 85:
        return "card-green"
    elif val >= 70:
        return "card-yellow"
    return "card-red"


def _global_accuracy(entries):
    """Overall weighted accuracy across all items (row-level)."""
    total_items = 0
    total_score = 0.0
    for e in entries:
        ad = e.get("accuracy_data", {})
        t = int(ad.get("total", 0) or 0)
        if t == 0:
            continue
        total_items += t
        total_score += sum(
            int(ad.get(lv, 0) or 0) * weights.get(lv, 0)
            for lv in ("full", "high", "medium", "low", "none")
        )
    return (total_score / total_items * 100) if total_items else 0


def _email_accuracy(entries):
    """Average of per-email weighted accuracy."""
    accs = []
    for e in entries:
        ad = e.get("accuracy_data", {})
        if int(ad.get("total", 0) or 0) > 0:
            accs.append(calc_weighted_accuracy(ad, weights))
    return sum(accs) / len(accs) if accs else 0


def _total_items(entries):
    return sum(int(e.get("items_count", 0) or e.get("accuracy_data", {}).get("total", 0) or 0) for e in entries)


def _total_cost(entries):
    return sum(float(e.get("cost_usd", 0) or 0) for e in entries)


def _total_time(entries):
    return sum(float(e.get("processing_time_seconds", 0) or 0) for e in entries)


def _total_files(entries):
    return sum(int(e.get("files_processed", 0) or 0) for e in entries)


def _confidence_totals(entries):
    """Sum confidence levels across all entries."""
    totals = {"full": 0, "high": 0, "medium": 0, "low": 0, "none": 0}
    for e in entries:
        ad = e.get("accuracy_data", {})
        for lv in totals:
            totals[lv] += int(ad.get(lv, 0) or 0)
    return totals


def _entries_by_day(entries):
    days = defaultdict(list)
    for e in entries:
        dt = _entry_ts(e)
        if dt != datetime.min:
            days[dt.strftime("%Y-%m-%d")].append(e)
    return dict(sorted(days.items(), reverse=True))


def _entry_message_id(entry):
    return str(entry.get("message_id") or "").strip()


def _entry_items(entry):
    return int(entry.get("items_count", 0) or entry.get("accuracy_data", {}).get("total", 0) or 0)


def _entry_files(entry):
    return int(entry.get("files_processed", 0) or 0)


def _unique_message_keys(entries, prefix):
    keys = set()
    for idx, entry in enumerate(entries):
        message_id = _entry_message_id(entry)
        if message_id:
            keys.add(message_id)
        else:
            keys.add(f"{prefix}:{idx}:{entry.get('timestamp') or entry.get('received') or ''}")
    return keys


def _is_no_draw_email_entry(entry):
    return (
        _entry_items(entry) == 0
        and _entry_files(entry) == 0
        and float(entry.get("cost_usd", 0) or 0) == 0.0
        and str(entry.get("type") or "").upper() != "RERUN"
    )


def _run_type_for_entry(entry, heavy_message_ids):
    run_type = str(entry.get("run_type") or "").strip().lower()
    if run_type in {"regular", "heavy"}:
        return run_type
    return "heavy" if _entry_message_id(entry) in heavy_message_ids else "regular"


def _build_operations_metrics(filtered_entries, mail_entries):
    event_entries = [entry for entry in filtered_entries if entry.get("event")]
    rerun_entries = [
        entry for entry in filtered_entries
        if str(entry.get("type") or "").upper() == "RERUN"
    ]
    heavy_marked_events = [entry for entry in event_entries if entry.get("event") == "heavy_email_skipped"]
    no_draw_events = [
        entry for entry in event_entries
        if entry.get("event") in {"no_draw_skipped", "no_attachment_skipped"}
    ]
    skip_sender_events = [entry for entry in event_entries if entry.get("event") == "skip_sender"]
    skip_category_events = [entry for entry in event_entries if entry.get("event") == "skip_category"]

    heavy_message_ids = _unique_message_keys(heavy_marked_events, "heavy")
    processable_entries = [
        entry for entry in mail_entries
        if str(entry.get("type") or "").upper() != "RERUN"
    ]
    no_draw_processed_entries = [entry for entry in processable_entries if _is_no_draw_email_entry(entry)]
    processed_entries = [entry for entry in processable_entries if not _is_no_draw_email_entry(entry)]
    heavy_processed_entries = [
        entry for entry in processed_entries
        if _run_type_for_entry(entry, heavy_message_ids) == "heavy"
    ]
    regular_processed_entries = [
        entry for entry in processed_entries
        if _run_type_for_entry(entry, heavy_message_ids) == "regular"
    ]

    no_draw_message_keys = _unique_message_keys(no_draw_processed_entries, "no-draw-processed")
    no_draw_message_keys.update(_unique_message_keys(no_draw_events, "no-draw-event"))
    skip_sender_keys = _unique_message_keys(skip_sender_events, "skip-sender")
    skip_category_keys = _unique_message_keys(skip_category_events, "skip-category")
    rerun_keys = _unique_message_keys(rerun_entries, "rerun")

    daily = defaultdict(lambda: {
        "regular_processed": 0,
        "heavy_processed": 0,
        "no_draw": 0,
        "skip_sender": 0,
        "skip_category": 0,
        "rerun": 0,
        "heavy_marked": 0,
    })

    for entry in regular_processed_entries:
        dt = _entry_ts(entry)
        if dt != datetime.min:
            daily[dt.strftime("%Y-%m-%d")]["regular_processed"] += 1
    for entry in heavy_processed_entries:
        dt = _entry_ts(entry)
        if dt != datetime.min:
            daily[dt.strftime("%Y-%m-%d")]["heavy_processed"] += 1
    for entry in no_draw_processed_entries:
        dt = _entry_ts(entry)
        if dt != datetime.min:
            daily[dt.strftime("%Y-%m-%d")]["no_draw"] += 1
    for entry in skip_sender_events:
        dt = _entry_ts(entry)
        if dt != datetime.min:
            daily[dt.strftime("%Y-%m-%d")]["skip_sender"] += 1
    for entry in skip_category_events:
        dt = _entry_ts(entry)
        if dt != datetime.min:
            daily[dt.strftime("%Y-%m-%d")]["skip_category"] += 1
    for entry in rerun_entries:
        dt = _entry_ts(entry)
        if dt != datetime.min:
            daily[dt.strftime("%Y-%m-%d")]["rerun"] += 1
    for entry in heavy_marked_events:
        dt = _entry_ts(entry)
        if dt != datetime.min:
            daily[dt.strftime("%Y-%m-%d")]["heavy_marked"] += 1

    all_unique_count = len(
        _unique_message_keys(processed_entries, "processed")
        | no_draw_message_keys
        | skip_sender_keys
        | skip_category_keys
        | rerun_keys
    )

    return {
        "processed_entries": processed_entries,
        "regular_processed_entries": regular_processed_entries,
        "heavy_processed_entries": heavy_processed_entries,
        "no_draw_processed_entries": no_draw_processed_entries,
        "rerun_entries": rerun_entries,
        "heavy_marked_events": heavy_marked_events,
        "skip_sender_events": skip_sender_events,
        "skip_category_events": skip_category_events,
        "processed_count": len(_unique_message_keys(processed_entries, "processed")),
        "regular_count": len(_unique_message_keys(regular_processed_entries, "regular")),
        "heavy_count": len(_unique_message_keys(heavy_processed_entries, "heavy-processed")),
        "no_draw_count": len(no_draw_message_keys),
        "skip_sender_count": len(skip_sender_keys),
        "skip_category_count": len(skip_category_keys),
        "skip_total_count": len(skip_sender_keys | skip_category_keys),
        "rerun_count": len(rerun_keys),
        "heavy_marked_count": len(_unique_message_keys(heavy_marked_events, "heavy-marked")),
        "all_unique_count": all_unique_count,
        "daily": dict(sorted(daily.items())),
    }


# ═══════  SUMMARY CARDS  ═══════
operations = _build_operations_metrics(filtered, email_entries)

total_emails = operations["all_unique_count"]
cost = _total_cost(email_entries)
items = _total_items(email_entries)
total_time_all = _total_time(email_entries)
row_acc = _global_accuracy(email_entries)
mail_acc = _email_accuracy(email_entries)
avg_time = total_time_all / total_emails if total_emails else 0
avg_time_row = total_time_all / items if items else 0
cost_per_email = cost / total_emails if total_emails else 0
cost_per_row = cost / items if items else 0

# ── Previous period for delta calculation ──
def _prev_period_entries(all_ents, period_key):
    """Get entries from the previous equivalent period for delta comparison."""
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    if period_key == "היום":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)
        end = now.replace(hour=0, minute=0, second=0, microsecond=0)
    elif period_key == "שבוע":
        start = now - timedelta(days=14)
        end = now - timedelta(days=7)
    elif period_key == "חודש":
        start = now - timedelta(days=60)
        end = now - timedelta(days=30)
    else:
        return []
    result = []
    for e in all_ents:
        ts = e.get("timestamp") or e.get("start_time") or ""
        try:
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00")).replace(tzinfo=None)
            if start <= dt < end:
                result.append(e)
        except (ValueError, AttributeError):
            pass
    return result

prev_period_entries = _prev_period_entries(all_entries, period_key)
prev_entries = [e for e in prev_period_entries if e.get("accuracy_data") or e.get("files_processed")]
prev_operations = _build_operations_metrics(prev_period_entries, prev_entries)
prev_emails = prev_operations["all_unique_count"]
prev_items = _total_items(prev_entries) if prev_entries else None
prev_mail_acc = _email_accuracy(prev_entries) if prev_entries else None
prev_row_acc = _global_accuracy(prev_entries) if prev_entries else None
prev_cost = _total_cost(prev_entries) if prev_entries else None
prev_total_time = _total_time(prev_entries) if prev_entries else None
prev_avg_time = (prev_total_time / prev_emails) if prev_emails else None
prev_avg_time_row = (prev_total_time / prev_items) if prev_items else None
prev_cost_per_email = (prev_cost / prev_emails) if prev_emails else None
prev_cost_per_row = (prev_cost / prev_items) if prev_items else None

def _delta(current, previous):
    if previous is None or previous == 0:
        return None
    return round(current - previous, 1)

def _delta_str(val):
    if val is None:
        return None
    return f"{val:+.1f}"

def _delta_str4(current, previous):
    """Delta string for cost values with 4 decimals."""
    if previous is None or previous == 0:
        return None
    d = current - previous
    return f"{d:+.4f}"

# Date range
dates = [_entry_ts(e) for e in email_entries if _entry_ts(e) != datetime.min]
if dates:
    first_date = min(dates).strftime("%d/%m")
    last_date = max(dates).strftime("%d/%m")
    period_label = f"{first_date} → {last_date}"
else:
    period_label = "—"

# Row 1: תקופה | כמויות | דיוק
r1 = st.columns([1.4, 1, 1, 1, 1])
r1[0].metric("📅 תקופה", period_label)
r1[1].metric("📬 מיילים", total_emails, delta=_delta(total_emails, prev_emails))
r1[2].metric("📋 שורות", items, delta=_delta(items, prev_items))
r1[3].metric("🎯 דיוק מיילים", f"{mail_acc:.1f}%", delta=_delta_str(_delta(mail_acc, prev_mail_acc)))
r1[4].metric("📐 דיוק שורות", f"{row_acc:.1f}%", delta=_delta_str(_delta(row_acc, prev_row_acc)))

# Row 2: זמנים | עלויות
r2 = st.columns(5)
r2[0].metric("⏱ זמן/מייל", f"{avg_time:.0f}s", delta=_delta_str(_delta(avg_time, prev_avg_time)), delta_color="inverse")
r2[1].metric("⏱ זמן/שורה", f"{avg_time_row:.1f}s", delta=_delta_str(_delta(avg_time_row, prev_avg_time_row)), delta_color="inverse")
r2[2].metric("💵 עלות/מייל", f"${cost_per_email:.3f}", delta=_delta_str4(cost_per_email, prev_cost_per_email), delta_color="inverse")
r2[3].metric("💵 עלות/שורה", f"${cost_per_row:.4f}", delta=_delta_str4(cost_per_row, prev_cost_per_row), delta_color="inverse")
r2[4].metric("💰 עלות כוללת", f"${cost:.2f}", delta=_delta_str(_delta(cost, prev_cost)), delta_color="inverse")

st.markdown("---")

if not email_entries:
    st.info("אין נתונים לתקופה שנבחרה")
    st.stop()


# ═══════  TABS  ═══════
tab_accuracy, tab_efficiency, tab_operations, tab_customers, tab_senders, tab_emails, tab_export = st.tabs([
    "🎯 דיוק", "⚡ יעילות", "🧭 תפעול", "👥 לקוחות", "📨 שולחים", "📧 הודעות אחרונות", "📊 ייצוא"
])


# ╔═══════════════════════════════════════════════╗
# ║  TAB 1: ACCURACY                              ║
# ╚═══════════════════════════════════════════════╝
with tab_accuracy:
    # --- Accuracy Weights Editor ---
    with st.expander("⚖️ משקלות דיוק (weights)", expanded=False):
        st.caption("ערכים אלו משפיעים על חישוב אחוזי הדיוק. שינויים יישמרו בקובץ .env.")
        wcols = st.columns(5)
        labels_w = {"full": "מלא (Full)", "high": "גבוה (High)", "medium": "בינוני (Medium)", "low": "נמוך (Low)", "none": "ללא (None)"}
        new_weights = {}
        for i, lv in enumerate(("full", "high", "medium", "low", "none")):
            with wcols[i]:
                new_weights[lv] = st.number_input(labels_w[lv], value=weights[lv], min_value=0.0, max_value=1.0, step=0.1, key=f"w_{lv}")
        if st.button("💾 שמור משקלות"):
            env_path = PROJECT_ROOT / ".env"
            env_lines = []
            if env_path.exists():
                env_lines = env_path.read_text(encoding="utf-8").splitlines()
            # Update or append weight lines
            env_keys = {"full": "ACCURACY_WEIGHT_FULL", "high": "ACCURACY_WEIGHT_HIGH",
                        "medium": "ACCURACY_WEIGHT_MEDIUM", "low": "ACCURACY_WEIGHT_LOW", "none": "ACCURACY_WEIGHT_NONE"}
            for lv, env_key in env_keys.items():
                found = False
                for j, el in enumerate(env_lines):
                    if el.strip().startswith(env_key + "="):
                        env_lines[j] = f"{env_key}={new_weights[lv]}"
                        found = True
                        break
                if not found:
                    env_lines.append(f"{env_key}={new_weights[lv]}")
            env_path.write_text("\n".join(env_lines) + "\n", encoding="utf-8")
            # Also update in-memory env vars so the page recalculates
            for lv, env_key in env_keys.items():
                os.environ[env_key] = str(new_weights[lv])
            st.toast("✅ משקלות נשמרו", icon="✅")
            st.rerun()

    # --- Confidence Distribution ---
    st.subheader("📊 התפלגות רמות ביטחון")
    conf = _confidence_totals(email_entries)
    total_conf = sum(conf.values())

    if total_conf > 0:
        conf_cols = st.columns(5)
        colors = {"full": "#10b981", "high": "#06b6d4", "medium": "#f59e0b", "low": "#ef4444", "none": "#6b7280"}
        labels = {"full": "מלא", "high": "גבוה", "medium": "בינוני", "low": "נמוך", "none": "ללא"}
        for i, lv in enumerate(("full", "high", "medium", "low", "none")):
            pct = conf[lv] / total_conf * 100
            c = colors[lv]
            conf_cols[i].markdown(
                f'<div style="text-align:center; background:rgba(13,17,23,0.6); '
                f'backdrop-filter:blur(10px); border:1px solid {c}22; border-radius:12px; '
                f'padding:16px 8px; position:relative; overflow:hidden;">'
                f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{c};"></div>'
                f'<div style="font-size:28px; font-weight:800; color:{c}; '
                f'text-shadow:0 0 15px {c}44;">{conf[lv]}</div>'
                f'<div style="font-size:11px; color:#94a3b8; margin-top:4px; '
                f'text-transform:uppercase; letter-spacing:1px;">{labels[lv]}</div>'
                f'<div style="font-size:18px; font-weight:700; color:{c}; margin-top:2px;">{pct:.0f}%</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        # Distribution bar with animated gradient
        bar_html = '<div style="height:24px; border-radius:12px; overflow:hidden; display:flex; '\
                   'margin-top:12px; box-shadow: 0 0 20px rgba(0,0,0,0.3); border:1px solid rgba(255,255,255,0.05);">'
        for lv in ("full", "high", "medium", "low", "none"):
            w = conf[lv] / total_conf * 100
            if w > 0:
                bar_html += f'<div style="width:{w}%; background:{colors[lv]}; height:100%; '\
                            f'box-shadow: inset 0 1px 0 rgba(255,255,255,0.15); '\
                            f'transition: width 0.5s ease;"></div>'
        bar_html += '</div>'
        st.markdown(bar_html, unsafe_allow_html=True)

    st.markdown("")

    # --- 3-Period Accuracy Grid ---
    st.subheader("📐 דיוק לפי תקופה")
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    today_entries = [e for e in email_entries if _entry_ts(e).date() == now.date()]
    month_entries = [e for e in email_entries if _entry_ts(e) >= now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)]

    grid_data = {
        "יומי": today_entries,
        "חודשי": month_entries,
        "כללי": email_entries,
    }

    gcols = st.columns(3)
    for i, (label, ents) in enumerate(grid_data.items()):
        with gcols[i]:
            st.markdown(f"**{label}**")
            ea = _email_accuracy(ents)
            ra = _global_accuracy(ents)
            st.markdown(f'📧 דיוק מיילים: <span class="{_acc_color(ea)}">{ea:.1f}%</span>', unsafe_allow_html=True)
            st.markdown(f'📋 דיוק שורות: <span class="{_acc_color(ra)}">{ra:.1f}%</span>', unsafe_allow_html=True)
            st.markdown(f'📬 מיילים: **{len(ents)}**')

    st.markdown("")

    # --- Accuracy Trend (last 14 days) ---
    st.subheader("📈 מגמת דיוק — 14 ימים אחרונים")
    days = _entries_by_day(email_entries)
    day_keys = list(days.keys())[:14]

    if day_keys:
        try:
            import plotly.graph_objects as go
            trend_dates = []
            trend_accs = []
            trend_counts = []
            for d in reversed(day_keys):
                trend_dates.append(d)
                trend_accs.append(_email_accuracy(days[d]))
                trend_counts.append(len(days[d]))

            fig = go.Figure()
            fig.add_trace(go.Bar(x=trend_dates, y=trend_counts, name="מיילים", marker_color="#0984e3", yaxis="y"))
            fig.add_trace(go.Scatter(x=trend_dates, y=trend_accs, name="דיוק %", mode="lines+markers",
                                     marker_color="#00b894", line_width=3, yaxis="y2"))
            fig.update_layout(
                yaxis=dict(title="מיילים", side="right"),
                yaxis2=dict(title="דיוק %", overlaying="y", side="left", range=[0, 105]),
                height=350, margin=dict(l=40, r=40, t=20, b=30),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                template="plotly_dark",
            )
            st.plotly_chart(fig, width='stretch')
        except ImportError:
            for d in day_keys:
                acc = _email_accuracy(days[d])
                n = len(days[d])
                bar = "█" * max(1, int(acc / 5))
                st.text(f"{d} | {bar} {acc:.0f}% ({n} מיילים)")

    # --- PL Override Stats ---
    total_pl = sum(int(e.get("pl_overrides", 0) or 0) for e in email_entries)
    emails_with_pl = sum(1 for e in email_entries if int(e.get("pl_overrides", 0) or 0) > 0)
    pl_pct = (total_pl / items * 100) if items else 0

    st.markdown("")
    st.subheader("🔧 PL Overrides")
    pl_cols = st.columns(3)
    pl_cols[0].metric("🔧 פריטים שתוקנו", total_pl)
    pl_cols[1].metric("📧 מיילים עם PL", emails_with_pl)
    pl_cols[2].metric("📊 אחוז תיקון", f"{pl_pct:.1f}%")

    # --- Top Errors ---
    error_counter = Counter()
    for e in email_entries:
        for err in (e.get("error_types") or []):
            error_counter[err] += 1

    if error_counter:
        st.markdown("")
        st.subheader("⚠️ שגיאות נפוצות")
        error_labels = {
            "missing_part_number": "חסר מספר פריט",
            "low_confidence": "ביטחון נמוך",
            "api_error": "שגיאת API",
            "timeout": "חריגת זמן",
        }
        for err, cnt in error_counter.most_common(5):
            label = error_labels.get(err, err)
            st.markdown(f"- **{label}**: {cnt}")


# ╔═══════════════════════════════════════════════╗
# ║  TAB 2: EFFICIENCY                            ║
# ╚═══════════════════════════════════════════════╝
with tab_efficiency:
    total_time_s = _total_time(email_entries)
    total_files_n = _total_files(email_entries)

    # Efficiency metrics
    st.subheader("⚡ מדדי יעילות")
    eff_cols = st.columns(4)
    eff_cols[0].metric("💰 עלות לפריט", f"${cost / items:.4f}" if items else "$0")
    eff_cols[1].metric("⏱ זמן לפריט", f"{total_time_s / items:.1f}s" if items else "0s")
    eff_cols[2].metric("📋 סה\"כ פריטים", items)
    eff_cols[3].metric("📄 סה\"כ קבצים", total_files_n)

    st.markdown("")

    # Items per mail stats
    st.subheader("📊 פריטים למייל")
    items_per_mail = [int(e.get("items_count", 0) or e.get("accuracy_data", {}).get("total", 0) or 0) for e in email_entries]
    items_per_mail_positive = [x for x in items_per_mail if x > 0]

    if items_per_mail_positive:
        import statistics
        ipm_cols = st.columns(6)
        ipm_cols[0].metric("📊 ממוצע", f"{statistics.mean(items_per_mail_positive):.1f}")
        ipm_cols[1].metric("📐 חציון", f"{statistics.median(items_per_mail_positive):.0f}")
        ipm_cols[2].metric("⬆️ מקסימום", max(items_per_mail_positive))
        ipm_cols[3].metric("⬇️ מינימום", min(items_per_mail_positive))
        ipm_cols[4].metric("📧 מיילים", len(items_per_mail_positive))
        ipm_cols[5].metric("📋 סה\"כ פריטים", sum(items_per_mail_positive))

        # Distribution chart
        buckets = {"1": 0, "2-3": 0, "4-5": 0, "6-10": 0, "11+": 0}
        for x in items_per_mail_positive:
            if x == 1:
                buckets["1"] += 1
            elif x <= 3:
                buckets["2-3"] += 1
            elif x <= 5:
                buckets["4-5"] += 1
            elif x <= 10:
                buckets["6-10"] += 1
            else:
                buckets["11+"] += 1

        try:
            import plotly.express as px
            fig = px.bar(
                x=list(buckets.keys()), y=list(buckets.values()),
                labels={"x": "פריטים למייל", "y": "מספר מיילים"},
                color_discrete_sequence=["#0984e3"],
            )
            fig.update_layout(height=280, template="plotly_dark", margin=dict(l=30, r=30, t=20, b=30))
            st.plotly_chart(fig, width='stretch')
        except ImportError:
            for label, cnt in buckets.items():
                bar = "█" * cnt
                st.text(f"{label}: {bar} {cnt}")

    st.markdown("")

    # Daily breakdown with cost
    st.subheader("📅 פירוט יומי")
    days = _entries_by_day(email_entries)
    day_keys = list(days.keys())[:14]

    if day_keys:
        try:
            import plotly.graph_objects as go
            d_dates, d_emails, d_costs = [], [], []
            for d in reversed(day_keys):
                d_dates.append(d)
                d_emails.append(len(days[d]))
                d_costs.append(_total_cost(days[d]))

            fig = go.Figure()
            fig.add_trace(go.Bar(x=d_dates, y=d_emails, name="מיילים", marker_color="#0984e3"))
            fig.add_trace(go.Scatter(x=d_dates, y=d_costs, name="עלות $", mode="lines+markers",
                                     marker_color="#e17055", line_width=3, yaxis="y2"))
            fig.update_layout(
                yaxis=dict(title="מיילים", side="right"),
                yaxis2=dict(title="$ עלות", overlaying="y", side="left"),
                height=320, margin=dict(l=40, r=40, t=20, b=30),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                template="plotly_dark",
            )
            st.plotly_chart(fig, width='stretch')
        except ImportError:
            pass

        # Table
        daily_rows = []
        for d in day_keys:
            ents = days[d]
            daily_rows.append({
                "תאריך": d,
                "מיילים": len(ents),
                "נשלחו": sum(1 for e in ents if e.get("sent")),
                "קבצים": _total_files(ents),
                "שורות": _total_items(ents),
                "עלות $": f"{_total_cost(ents):.3f}",
                "זמן (דקות)": f"{_total_time(ents) / 60:.1f}",
                "דיוק %": f"{_email_accuracy(ents):.1f}",
            })
        st.dataframe(daily_rows, width='stretch')


# ╔═══════════════════════════════════════════════╗
# ║  TAB 3: OPERATIONS                            ║
# ╚═══════════════════════════════════════════════╝
with tab_operations:
    st.subheader("🧭 סטטוס תפעולי")
    st.caption("דילוגים נספרים לפי אירועי JSONL. זה לא תמיד זהה להפרש חשבוני, כי אותו מייל יכול להופיע בכמה סטטוסים לאורך הזמן.")

    arithmetic_gap = max(
        0,
        operations["all_unique_count"]
        - (operations["processed_count"] + operations["no_draw_count"] + operations["rerun_count"]),
    )

    op0 = st.columns(3)
    op0[0].metric("📨 סה\"כ מיילים ייחודיים", operations["all_unique_count"])
    op0[1].metric("🔄 RERUN", operations["rerun_count"])
    op0[2].metric("➗ הפרש חשבוני", arithmetic_gap)

    op1 = st.columns(5)
    op1[0].metric("✅ מיילים שעובדו", operations["processed_count"])
    op1[1].metric("📬 רגילים", operations["regular_count"])
    op1[2].metric("🏋️ כבדים", operations["heavy_count"])
    op1[3].metric("📭 NO DRAW", operations["no_draw_count"])
    op1[4].metric("🚫 דילוגים", operations["skip_total_count"])

    regular_entries = operations["regular_processed_entries"]
    heavy_entries = operations["heavy_processed_entries"]
    regular_rows = _total_items(regular_entries)
    heavy_rows = _total_items(heavy_entries)
    regular_avg_rows = (regular_rows / operations["regular_count"]) if operations["regular_count"] else 0
    heavy_avg_rows = (heavy_rows / operations["heavy_count"]) if operations["heavy_count"] else 0
    regular_avg_time = (_total_time(regular_entries) / len(regular_entries)) if regular_entries else 0
    heavy_avg_time = (_total_time(heavy_entries) / len(heavy_entries)) if heavy_entries else 0

    op2 = st.columns(6)
    op2[0].metric("📋 שורות רגיל", regular_rows)
    op2[1].metric("📋 שורות כבד", heavy_rows)
    op2[2].metric("📐 ממוצע שורות רגיל", f"{regular_avg_rows:.1f}")
    op2[3].metric("📐 ממוצע שורות כבד", f"{heavy_avg_rows:.1f}")
    op2[4].metric("⏱ ממוצע רגיל", f"{regular_avg_time:.0f}s")
    op2[5].metric("⏱ ממוצע כבד", f"{heavy_avg_time:.0f}s")

    st.markdown("")

    try:
        import plotly.graph_objects as go

        status_labels = ["עובדו רגיל", "עובדו כבד", "NO DRAW", "RERUN", "דילוגים"]
        status_values = [
            operations["regular_count"],
            operations["heavy_count"],
            operations["no_draw_count"],
            operations["rerun_count"],
            operations["skip_total_count"],
        ]
        status_colors = ["#10b981", "#0ea5e9", "#f59e0b", "#eab308", "#ef4444"]

        fig_status = go.Figure(
            data=[go.Bar(x=status_labels, y=status_values, marker_color=status_colors, text=status_values, textposition="outside")]
        )
        fig_status.update_layout(height=320, margin=dict(l=30, r=30, t=20, b=30), template="plotly_dark")
        st.plotly_chart(fig_status, width='stretch')

        chart_cols = st.columns(2)
        with chart_cols[0]:
            fig_rows = go.Figure(
                data=[go.Pie(labels=["שורות רגיל", "שורות כבד"], values=[regular_rows, heavy_rows], hole=0.55,
                             marker=dict(colors=["#22c55e", "#38bdf8"]))]
            )
            fig_rows.update_layout(height=320, margin=dict(l=20, r=20, t=20, b=20), template="plotly_dark")
            st.plotly_chart(fig_rows, width='stretch')

        with chart_cols[1]:
            time_fig = go.Figure()
            if regular_entries:
                time_fig.add_trace(go.Box(
                    y=[float(entry.get("processing_time_minutes", 0) or 0) for entry in regular_entries],
                    name="רגיל",
                    marker_color="#22c55e",
                    boxmean=True,
                ))
            if heavy_entries:
                time_fig.add_trace(go.Box(
                    y=[float(entry.get("processing_time_minutes", 0) or 0) for entry in heavy_entries],
                    name="כבד",
                    marker_color="#38bdf8",
                    boxmean=True,
                ))
            time_fig.update_layout(height=320, margin=dict(l=30, r=30, t=20, b=30), template="plotly_dark", yaxis_title="דקות")
            st.plotly_chart(time_fig, width='stretch')

        if operations["daily"]:
            st.markdown("")
            st.subheader("📈 מגמת סטטוסים")
            day_labels = list(operations["daily"].keys())[-14:]
            fig_daily = go.Figure()
            fig_daily.add_trace(go.Bar(
                x=day_labels,
                y=[operations["daily"][day]["regular_processed"] for day in day_labels],
                name="רגיל",
                marker_color="#22c55e",
            ))
            fig_daily.add_trace(go.Bar(
                x=day_labels,
                y=[operations["daily"][day]["heavy_processed"] for day in day_labels],
                name="כבד",
                marker_color="#38bdf8",
            ))
            fig_daily.add_trace(go.Bar(
                x=day_labels,
                y=[operations["daily"][day]["no_draw"] for day in day_labels],
                name="NO DRAW",
                marker_color="#f59e0b",
            ))
            fig_daily.add_trace(go.Bar(
                x=day_labels,
                y=[operations["daily"][day]["rerun"] for day in day_labels],
                name="RERUN",
                marker_color="#eab308",
            ))
            fig_daily.add_trace(go.Scatter(
                x=day_labels,
                y=[
                    operations["daily"][day]["skip_sender"] + operations["daily"][day]["skip_category"]
                    for day in day_labels
                ],
                name="דילוגים",
                mode="lines+markers",
                marker_color="#ef4444",
                yaxis="y2",
            ))
            fig_daily.update_layout(
                barmode="stack",
                height=360,
                margin=dict(l=30, r=30, t=20, b=30),
                template="plotly_dark",
                yaxis=dict(title="כמות מיילים"),
                yaxis2=dict(title="דילוגים", overlaying="y", side="left"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig_daily, width='stretch')
    except ImportError:
        pass

    st.markdown("")
    st.subheader("📋 סיכום תפעולי")
    st.dataframe([
        {
            "קטגוריה": "עובדו רגיל",
            "מיילים": operations["regular_count"],
            "שורות": regular_rows,
            "זמן ממוצע (שניות)": round(regular_avg_time, 1),
        },
        {
            "קטגוריה": "עובדו כבד",
            "מיילים": operations["heavy_count"],
            "שורות": heavy_rows,
            "זמן ממוצע (שניות)": round(heavy_avg_time, 1),
        },
        {
            "קטגוריה": "NO DRAW",
            "מיילים": operations["no_draw_count"],
            "שורות": 0,
            "זמן ממוצע (שניות)": 0,
        },
        {
            "קטגוריה": "RERUN",
            "מיילים": operations["rerun_count"],
            "שורות": 0,
            "זמן ממוצע (שניות)": 0,
        },
        {
            "קטגוריה": "דילוגים",
            "מיילים": operations["skip_total_count"],
            "שורות": 0,
            "זמן ממוצע (שניות)": 0,
        },
    ], width='stretch')


# ╔═══════════════════════════════════════════════╗
# ║  TAB 3: CUSTOMERS                             ║
# ╚═══════════════════════════════════════════════╝
with tab_customers:
    st.subheader("👥 סטטיסטיקת לקוחות — Top 10")

    customer_data = defaultdict(lambda: {
        "emails": 0, "items": 0, "cost": 0.0, "time": 0.0,
        "sent": 0, "accuracy_data": {"full": 0, "high": 0, "medium": 0, "low": 0, "none": 0, "total": 0},
        "senders": set(),
    })

    for e in email_entries:
        customers = e.get("customers") or []
        if not customers:
            customers = ["לא ידוע"]
        for cust in customers:
            cd = customer_data[cust]
            cd["emails"] += 1
            cd["items"] += int(e.get("items_count", 0) or 0)
            cd["cost"] += float(e.get("cost_usd", 0) or 0)
            cd["time"] += float(e.get("processing_time_seconds", 0) or 0)
            cd["sent"] += 1 if e.get("sent") else 0
            ad = e.get("accuracy_data", {})
            for lv in ("full", "high", "medium", "low", "none", "total"):
                cd["accuracy_data"][lv] += int(ad.get(lv, 0) or 0)
            cd["senders"].add((e.get("sender") or "").lower())

    # Sort by items descending
    top_customers = sorted(customer_data.items(), key=lambda x: x[1]["items"], reverse=True)[:10]

    if top_customers:
        cust_rows = []
        for cust_name, cd in top_customers:
            acc = calc_weighted_accuracy(cd["accuracy_data"], weights)
            items_pct = (cd["items"] / items * 100) if items else 0
            emails_pct = (cd["emails"] / total_emails * 100) if total_emails else 0
            success_rate = (cd["sent"] / cd["emails"] * 100) if cd["emails"] else 0
            cust_rows.append({
                "לקוח": cust_name,
                "מיילים": cd["emails"],
                "% מיילים": f"{emails_pct:.1f}",
                "שורות": cd["items"],
                "% שורות": f"{items_pct:.1f}",
                "דיוק %": f"{acc:.1f}",
                "עלות $": f"{cd['cost']:.3f}",
                "עלות/פריט $": f"{cd['cost'] / cd['items']:.4f}" if cd["items"] else "—",
                "זמן (דקות)": f"{cd['time'] / 60:.1f}",
                "שולחים": len(cd["senders"] - {""}),
                "הצלחה %": f"{success_rate:.0f}",
            })
        st.dataframe(cust_rows, width='stretch', height=420)

        # Customer accuracy chart
        try:
            import plotly.express as px
            chart_data = [{"לקוח": r["לקוח"], "דיוק": float(r["דיוק %"]), "שורות": r["שורות"]} for r in cust_rows]
            fig = px.bar(chart_data, x="לקוח", y="דיוק", color="שורות", color_continuous_scale="teal",
                         labels={"דיוק": "דיוק %"})
            fig.update_layout(height=320, template="plotly_dark", margin=dict(l=30, r=30, t=20, b=30))
            st.plotly_chart(fig, width='stretch')
        except ImportError:
            pass
    else:
        st.info("אין נתוני לקוחות")


# ╔═══════════════════════════════════════════════╗
# ║  TAB 4: SENDERS                               ║
# ╚═══════════════════════════════════════════════╝
with tab_senders:
    st.subheader("📨 שולחים מובילים — Top 10")

    sender_data = defaultdict(lambda: {
        "emails": 0, "items": 0, "cost": 0.0, "time": 0.0,
        "sent": 0, "customers": set(),
        "accuracy_data": {"full": 0, "high": 0, "medium": 0, "low": 0, "none": 0, "total": 0},
    })

    for e in email_entries:
        sender = (e.get("sender") or "").lower().strip()
        if not sender:
            continue
        sd = sender_data[sender]
        sd["emails"] += 1
        sd["items"] += int(e.get("items_count", 0) or 0)
        sd["cost"] += float(e.get("cost_usd", 0) or 0)
        sd["time"] += float(e.get("processing_time_seconds", 0) or 0)
        sd["sent"] += 1 if e.get("sent") else 0
        for c in (e.get("customers") or []):
            sd["customers"].add(c)
        ad = e.get("accuracy_data", {})
        for lv in ("full", "high", "medium", "low", "none", "total"):
            sd["accuracy_data"][lv] += int(ad.get(lv, 0) or 0)

    top_senders = sorted(sender_data.items(), key=lambda x: x[1]["emails"], reverse=True)[:10]

    if top_senders:
        sender_rows = []
        for sender_name, sd in top_senders:
            acc = calc_weighted_accuracy(sd["accuracy_data"], weights)
            success_rate = (sd["sent"] / sd["emails"] * 100) if sd["emails"] else 0
            sender_rows.append({
                "שולח": sender_name,
                "לקוחות": ", ".join(sorted(sd["customers"])) if sd["customers"] else "—",
                "מיילים": sd["emails"],
                "קבצים": sd["items"],
                "שורות": sd["items"],
                "דיוק %": f"{acc:.1f}",
                "עלות $": f"{sd['cost']:.3f}",
                "עלות/פריט $": f"{sd['cost'] / sd['items']:.4f}" if sd["items"] else "—",
                "זמן (דקות)": f"{sd['time'] / 60:.1f}",
                "הצלחה %": f"{success_rate:.0f}",
            })
        st.dataframe(sender_rows, width='stretch', height=420)
    else:
        st.info("אין נתוני שולחים")


# ╔═══════════════════════════════════════════════╗
# ║  TAB 5: RECENT EMAILS                         ║
# ╚═══════════════════════════════════════════════╝
with tab_emails:
    st.subheader("📧 הודעות אחרונות — 100")

    import pandas as pd
    table_data = []
    entry_ids = []
    for e in email_entries[:100]:
        ad = e.get("accuracy_data", {})
        acc = calc_weighted_accuracy(ad, weights) if int(ad.get("total", 0) or 0) > 0 else 0
        time_s = float(e.get("processing_time_seconds", 0) or 0)
        entry_ids.append(e.get("id", ""))
        table_data.append({
            "תאריך": (e.get("received") or e.get("timestamp") or "")[:16],
            "שולח": (e.get("sender") or "")[:32],
            "לקוחות": ", ".join(e.get("customers") or []),
            "קבצים": int(e.get("files_processed", 0) or 0),
            "שורות": int(e.get("items_count", 0) or 0),
            "עלות $": f"{float(e.get('cost_usd', 0) or 0):.3f}",
            "זמן": f"{time_s:.0f}s",
            "דיוק %": f"{acc:.1f}",
            "סטטוס": "✅" if e.get("sent") else "❌",
            "✓ אימות": bool(e.get("human_verified", True)),
        })

    if table_data:
        df = pd.DataFrame(table_data)
        edited = st.data_editor(
            df, height=500, use_container_width=True,
            disabled=["תאריך", "שולח", "לקוחות", "קבצים", "שורות", "עלות $", "זמן", "דיוק %", "סטטוס"],
            column_config={
                "✓ אימות": st.column_config.CheckboxColumn("✓ אימות", help="הסר סימון אם המייל כושל", default=True),
            },
            key="emails_editor",
        )
        # Save verification changes
        if edited is not None:
            for idx in range(len(table_data)):
                orig_val = table_data[idx]["✓ אימות"]
                new_val = bool(edited.iloc[idx]["✓ אימות"])
                if new_val != orig_val and entry_ids[idx]:
                    save_entry_field(entry_ids[idx], "human_verified", new_val)
    else:
        st.info("אין נתונים")


# ╔═══════════════════════════════════════════════╗
# ║  TAB 6: EXPORT                                 ║
# ╚═══════════════════════════════════════════════╝
with tab_export:
    st.subheader("📊 ייצוא לאקסל")
    st.markdown("ייצוא כל הנתונים לקובץ Excel עם 5 גיליונות: סיכום כללי, נתוני מיילים, סיכום יומי, לקוחות, שולחים.")

    if st.button("📥 הורד אקסל", width='stretch'):
        try:
            wb = build_workbook_dashboard(
                email_entries,
                run_label=f"dashboard:{period_key}",
                all_entries=filtered,
            )
            excel_bytes = workbook_to_bytes(wb)

            st.download_button(
                label="💾 שמור קובץ Excel",
                data=excel_bytes,
                file_name=f"DrawingAI_Stats_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("✅ הקובץ מוכן להורדה!")

        except ImportError as exc:
            st.error(f"חסרה חבילה: {exc}. התקן: pip install openpyxl")
        except Exception as exc:
            st.error(f"שגיאה בייצוא: {exc}")

    st.markdown("---")
    st.subheader("🗑️ איפוס סטטיסטיקה")
    st.warning("פעולה זו תגבה את הלוג הנוכחי ותאפס את הסטטיסטיקה.")
    if "confirm_reset_stats" not in st.session_state:
        st.session_state.confirm_reset_stats = False
    if st.session_state.confirm_reset_stats:
        reset_stats_btn = st.button("⚠️ לחץ שוב לאישור איפוס", type="primary")
    else:
        reset_stats_btn = st.button("🗑️ אפס סטטיסטיקה", type="secondary")
    if reset_stats_btn:
        if st.session_state.confirm_reset_stats:
            import shutil
            src = LOG_DIR / "automation_log.jsonl"
            if src.exists():
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup = LOG_DIR / f"automation_log_backup_{ts}.jsonl"
                shutil.copy2(src, backup)
                src.write_text("", encoding="utf-8")
                st.session_state.confirm_reset_stats = False
                st.success(f"✅ גיבוי נשמר ב-{backup.name} — הסטטיסטיקה אופסה")
                st.rerun()
        else:
            st.session_state.confirm_reset_stats = True
            st.rerun()

