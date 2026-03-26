# DrawingAI Pro — Data Flow

> עדכון אחרון: 25/03/2026 — **כולל Streamlit data flows**

## 📊 זרימת נתונים כללית

```
  Email Inbox (Graph API / EWS)
       │
       ▼
  automation_runner.py ──→ scan_folder()
       │                       │
       │  ┌────────────────────┘
       │  │
       ▼  ▼
  customer_extractor_v3_dual.py
       │
       ├─→ classify_file_type() ──→ DRAWING / PO / QUOTE / PL / ...
       │
       ├─→ extract_drawing_data() ──→ Pipeline (Stages 0-4)
       │        │
       │        ├─→ Stage 0: Layout
       │        ├─→ Stage 0.5: Rotation
       │        ├─→ Stage 1: Basic Info (P.N., customer)
       │        ├─→ Stage 2: Processes (material, coating, BOM)
       │        ├─→ Stage 3: Notes
       │        ├─→ Stage 4: Area
       │        ├─→ vote_best_pn() — 3-way voting
       │        ├─→ run_pn_sanity_checks() — checks A-D
       │        └─→ calculate_confidence() — full/high/medium/low/none
       │
       ├─→ match_quantities_to_drawings()
       ├─→ override_pn_from_email()
       ├─→ merge_descriptions() — Stage 9 (o4-mini)
       │        └─→ lookup_color_prices() — BOM/COLORS.xlsx
       ├─→ validate_inserts_hardware()
       ├─→ enrich_inserts_with_prices() — BOM/INSERTS.xlsx
       ├─→ rename_files_by_classification()
       │
       ├─→ _save_results_to_excel() ──→ Excel report
       ├─→ _save_text_summary() ──→ B2B text files
       └─→ send_email() ──→ Outlook
              │
              ▼
       automation_log.jsonl ──→ JSONL entry per email
```

---

## 📧 זרימת מייל (automation_runner.py)

```
  1. list_messages(received_after, max_messages)
  2. Filter: skip_senders, skip_categories, processed IDs
  3. _count_drawing_files(message_dir)
       ├─→ ≤ max_files_per_email → process normally
       └─→ > max_files_per_email → mark "AI HEAVY", skip
  4. scan_folder(message_dir, config)
  5. Log entry → automation_log.jsonl
  6. Send B2B email → recipient
  7. Mark processed → automation_state.json
  8. Set category in Outlook → ensure_category + replace_category
```

### Heavy Email Flow
```
  run_heavy() → _run_once_internal(heavy_only=True)
       → only processes messages with "AI HEAVY" category
       → no max_files_per_email limit
       → removes "AI HEAVY" category after processing
```

---

## 📋 result_dict Schema

```python
{
    "filename": str,
    "file_type": "DRAWING" | "PURCHASE_ORDER" | "PARTS_LIST" | ...,
    "customer_name": str,
    "part_number": str,
    "drawing_number": str,
    "revision": str,
    "material": str,
    "surface_coating": str,
    "color_painting": str,
    "geometric_area": float | None,
    "dimensions_raw": str,
    "processes_list": list[str],
    "bom_items": list[dict],
    "notes_text": str,
    "quantity": int,
    "confidence_level": "full" | "high" | "medium" | "low" | "none",
    "merged_description": str,          # Stage 9 output
    "merged_specs": str,                # Stage 9 output
    "merged_highlights": str,           # Stage 9 output
    "color_price": float | None,        # from COLORS.xlsx
    "hardware_items": list[dict],       # from BOM + INSERTS.xlsx
    "stage_costs": dict,
    "total_cost_usd": float,
    "processing_time_seconds": float,
}
```

---

## 📊 automation_log.jsonl Entry Schema

```json
{
    "id": "auto_YYYYMMDDHHMMSS_sender",
    "timestamp": "2026-03-12T10:30:00Z",
    "received": "2026-03-12T10:25:00Z",
    "sender": "user@company.com",
    "customers": ["CUSTOMER_A"],
    "files_processed": 3,
    "items_count": 5,
    "accuracy_data": {
        "full": 3, "high": 1, "medium": 1, "low": 0, "none": 0, "total": 5
    },
    "cost_usd": 0.103,
    "processing_time_seconds": 262,
    "sent": true,
    "pl_overrides": 1,
    "error_types": [],
    "human_verified": false
}
```

---

## 🖥️ ★ Streamlit Data Flows

### Automation Page — Data Flow
```
  Browser
    │
    ▼
  session_state: runner, is_running, status_msg, log_lines, confirm_reset
    │
    ├─→ config_manager.load_config() ──→ automation_config.json ──→ form fields
    │
    ├─→ _header_status_fragment() [every 5s]
    │     ├─→ runner.get_run_status() ──→ status bar
    │     ├─→ detect_active_run() ──→ heavy status
    │     └─→ load_log_entries() → filter_by_period("today") ──→ cost display
    │
    ├─→ _live_log_fragment() [every 5s]
    │     └─→ read_log_tail() ──→ status_log.txt ──→ HTML container
    │
    ├─→ Save button ──→ _gather_config() ──→ save_config()
    ├─→ Run Once ──→ runner.run_once() (thread) ──→ st.status() progress
    ├─→ Run Heavy ──→ runner.run_heavy() (thread) ──→ st.status() progress
    ├─→ Start/Stop ──→ runner.start()/stop()
    └─→ Reset ──→ confirm_reset → reset_state()
```

### Dashboard Page — Data Flow
```
  load_log_entries() ──→ all JSONL files (deduplicated)
    │
    ├─→ filter_by_period(entries, period) ──→ email_entries
    │
    ├─→ KPI Cards (10):
    │     ├─→ _total_items/cost/time() ──→ current values
    │     └─→ _prev_period_entries() ──→ delta comparison
    │
    ├─→ Tab: Accuracy
    │     ├─→ _confidence_totals() ──→ distribution bar
    │     ├─→ _global_accuracy() / _email_accuracy() ──→ period grid
    │     ├─→ _entries_by_day() ──→ Plotly 14-day trend
    │     └─→ Weights Editor ──→ .env file (ACCURACY_WEIGHT_*)
    │
    ├─→ Tab: Efficiency
    │     └─→ statistics + Plotly distribution + daily breakdown
    │
    ├─→ Tab: Customers/Senders
    │     └─→ Top 10 tables + Plotly charts
    │
    ├─→ Tab: Recent Emails
    │     └─→ st.data_editor ──→ ✓ verification ──→ save_entry_field() ──→ JSONL
    │
    └─→ Tab: Export
          ├─→ openpyxl ──→ Excel (5 sheets) ──→ st.download_button
          └─→ Reset Stats ──→ confirm → backup + clear JSONL
```

### Accuracy Weights Flow
```
  .env file
    │ ACCURACY_WEIGHT_FULL=1.0
    │ ACCURACY_WEIGHT_HIGH=1.0
    │ ACCURACY_WEIGHT_MEDIUM=0.8
    │ ACCURACY_WEIGHT_LOW=0.5
    │ ACCURACY_WEIGHT_NONE=0.0
    │
    ├─→ get_accuracy_weights() ──→ os.getenv()
    │     └─→ calc_weighted_accuracy() ──→ per-entry score
    │
    └─→ Dashboard Weights Editor
          └─→ Save ──→ write .env + os.environ update ──→ st.rerun()
```

### Human Verification Flow
```
  Dashboard "Recent Emails" tab
    │
    ├─→ Load entries[:100] ──→ pd.DataFrame
    ├─→ st.data_editor (checkbox column "✓ אימות")
    └─→ On change ──→ save_entry_field(entry_id, "human_verified", bool)
                        └─→ Rewrite JSONL line in-place
```

---

## 🔄 OCR Fallback Chain

```
  Stage 1 extract_basic_info
    │
    ├─→ Vision API (full image + title block)
    │     └─→ Success? → done
    │
    ├─→ MultiOCREngine
    │     ├─→ pytesseract (Hebrew + English)
    │     ├─→ Azure Vision API OCR
    │     └─→ combined + deduplicated
    │
    └─→ extract_stage1_with_retry (higher DPI, bigger crop)
```

---

## 📄 Excel Output Sheets

| Sheet | Content |
|-------|---------|
| תוצאות | Main results per drawing |
| סיווג קבצים | File classification report |
| Parts List | PL items + associated drawings |
| BOM | Hardware + insert prices |
| סיכום | Summary statistics |

---

## 📤 B2B Output Format

```
Field 1:  מק"ט (P.N.)
Field 2:  גרסה (Revision)
Field 3:  שם לקוח
Field 4:  כמות
Field 5:  חומר
Field 6:  שטח (m²)
Field 7:  ציפוי פנים
Field 8:  צביעה חיצונית
Field 9:  מחיר צבע
Field 10: הערות
Field 11: merged_description (תיאור מורחב — Stage 9)
```
