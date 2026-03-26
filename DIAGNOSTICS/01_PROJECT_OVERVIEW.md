# DrawingAI Pro — סקירת פרויקט

> עדכון אחרון: 25/03/2026 — **ניקוי קבצים שלא בשימוש, עדכון ספירות ומבנה**

## 📌 תיאור כללי

**DrawingAI Pro** הוא מערכת לניתוח אוטומטי של שרטוטים הנדסיים באמצעות Azure OpenAI Vision API.
המערכת מזהה סוגי קבצים, מחלצת מידע הנדסי (לקוח, מק"ט, חומר, תהליכים, ציפויים, קשיחים מ-BOM),
ממזגת תיאורים חכמים ב-Stage 9 (o4-mini), מבצעת חיפוש מחירי צבע וקשיחים מקטלוג,
ומפיקה דוחות Excel ו-B2B.

### נקודות כניסה

| קובץ | תפקיד | שורות |
|-------|--------|------:|
| `Run_Web.bat` | מפעיל Streamlit Web UI | — |
| `Run_GUI.bat` | מפעיל את ה-GUI הראשי (Tkinter Legacy) | — |
| `streamlit_app/app.py` | ★ Streamlit entry point — ממשק Web ראשי | 42 |
| `customer_extractor_gui.py` | ממשק גרפי (Tkinter Legacy) | 1,303 |
| `customer_extractor_v3_dual.py` | מנוע הליבה (pipeline) | 2,127 |
| `main.py` | נקודת כניסה CLI | 50 |
| `automation_runner.py` | הרצה אוטומטית מחזורית + heavy email | 1,731 |
| `automation_main.py` | Tkinter entry point | 66 |

---

## 🏗️ ארכיטקטורה — מבנה תקיות

```
AI DRAW/
├── customer_extractor_v3_dual.py    ← מנוע ליבה (pipeline ראשי) — 2,127 שורות
├── customer_extractor_gui.py        ← GUI ראשי (Tkinter Legacy) — 1,303 שורות
├── automation_runner.py             ← runner אוטומטי + heavy email — 1,731 שורות
├── process_analysis.py              ← סטטיסטיקות תהליכים/חומרים — 403 שורות
├── main.py                          ← CLI entry point — 50 שורות
├── automation_main.py               ← Tkinter entry point — 66 שורות
│
├── streamlit_app/                   ← ★ Streamlit Web UI — 2,679 שורות
│   ├── app.py                       ← 42 שורות — entry point + page config
│   ├── brand.py                     ← 492 שורות — לוגו, CSS, brand (Green Coat/Algat)
│   ├── backend/
│   │   ├── config_manager.py        ← 88 שורות — R/W automation_config.json
│   │   ├── email_helpers.py         ← 133 שורות — חיבור תיבות משותפות + folders
│   │   ├── excel_report_builder.py  ← 539 שורות — Excel report generation (dashboard export)
│   │   ├── log_reader.py            ← 373 שורות — JSONL log + live log + detection
│   │   ├── report_exporter.py       ← 37 שורות — report export helper
│   │   └── runner_bridge.py         ← 306 שורות — thread-safe AutomationRunner wrapper
│   ├── components/
│   └── pages/
│       ├── 1_🚀_Automation.py       ← 715 שורות — פאנל אוטומציה מלא (4 tabs)
│       ├── 2_📊_Dashboard.py        ← 936 שורות — דשבורד סטטיסטיקה (6 tabs)
│       └── 3_📧_Email.py            ← 79 שורות — ניהול מייל
│
├── src/                             ← מודולים מחולצים — 14,495 שורות
│   ├── core/                        ← קבועים, הגדרות, exceptions
│   │   ├── constants.py             ← 90 שורות — קבועים משותפים (HUB)
│   │   ├── config.py                ← 181 שורות — Config classes
│   │   ├── cost_tracker.py          ← 78 שורות — מעקב עלויות
│   │   └── exceptions.py            ← 65 שורות — exception hierarchy
│   │
│   ├── services/
│   │   ├── ai/
│   │   │   ├── model_runtime.py     ← 254 שורות — ModelRuntimeConfig + endpoint helpers
│   │   │   └── vision_api.py        ← 279 שורות — Vision API + GPT-4o fallback
│   │   ├── image/
│   │   │   └── processing.py        ← 812 שורות — סיבוב, downsample, quality
│   │   ├── extraction/              ← חילוץ מידע — 16 מודולים
│   │   │   ├── stages_generic.py    ← 437 שורות
│   │   │   ├── stages_rafael.py     ← 314 שורות
│   │   │   ├── stages_iai.py        ← 310 שורות
│   │   │   ├── stage9_merge.py      ← 373 שורות — Stage 9: o4-mini merge
│   │   │   ├── color_price_lookup.py← 324 שורות
│   │   │   ├── insert_price_lookup.py←172 שורות
│   │   │   ├── insert_validator.py  ← 133 שורות
│   │   │   ├── ocr_engine.py        ← 554 שורות
│   │   │   ├── filename_utils.py    ← 594 שורות
│   │   │   ├── document_reader.py   ← 783 שורות
│   │   │   ├── drawing_pipeline.py  ← 775 שורות
│   │   │   ├── pn_voting.py         ← 238 שורות
│   │   │   ├── sanity_checks.py     ← 555 שורות
│   │   │   ├── post_processing.py   ← 135 שורות
│   │   │   └── quantity_matcher.py  ← 399 שורות
│   │   ├── file/
│   │   │   ├── file_utils.py        ← 708 שורות
│   │   │   ├── classifier.py        ← 339 שורות
│   │   │   └── file_renamer.py      ← 89 שורות
│   │   ├── reporting/
│   │   │   ├── b2b_export.py        ← 258 שורות
│   │   │   ├── pl_generator.py      ← 942 שורות
│   │   │   └── excel_export.py      ← 610 שורות
│   │   └── email/
│   │       ├── shared_mailbox.py    ← 593 שורות — EWS
│   │       ├── graph_mailbox.py     ← 916 שורות — Graph API
│   │       ├── graph_helper.py      ← 528 שורות
│   │       ├── graph_auth.py        ← 261 שורות
│   │       └── factory.py           ← 168 שורות
│   ├── models/
│   │   ├── drawing.py               ← 180 שורות
│   │   └── enums.py                 ← 48 שורות
│   └── utils/
│       ├── logger.py                ← 141 שורות
│       └── prompt_loader.py         ← 39 שורות
│
├── prompts/                         ← 15 פרומפטים ל-AI
├── tests/                           ← 25 קבצי בדיקות
├── BOM/                             ← COLORS.xlsx, INSERTS.xlsx
├── deploy/                          ← install_server.ps1, register_service.ps1, UPDATE.bat
├── .streamlit/                      ← config.toml (theme, server, client)
└── logs/                            ← קבצי לוג
```

---

## 📊 סיכום שורות קוד

| תחום | ערך |
|-------|──────:|
| **סה"כ Python files** | **102** |
| פרומפטים AI | 15 קבצים |
| בדיקות | 25 קבצים |

---

## 🖥️ ★ Streamlit Web UI — ארכיטקטורה

### מבנה שכבות

```
┌──────────────────────────────────┐
│  Browser (localhost:8501)         │
├──────────────────┬───────────────┤
│  Pages (3)       │  Brand/CSS    │
│  🚀 Automation   │  brand.py     │
│  📊 Dashboard    │  (RTL, dark,  │
│  📧 Email        │  orange/green)│
├──────────────────┴───────────────┤
│  Backend Layer                    │
│  config_manager │ runner_bridge  │
│  log_reader     │ email_helpers  │
├──────────────────────────────────┤
│  Core Engine                      │
│  automation_runner.py             │
│  customer_extractor_v3_dual.py    │
└──────────────────────────────────┘
```

### 🚀 Automation Page (715 שורות)
- **Header**: auto-refresh `@st.fragment(run_every=5)` — סטטוס רגיל + כבדים + עלות יומית
- **7 כפתורים**: שמור, בדוק, Run Once, Run Heavy, התחל, עצור, Reset (confirmation)
- **4 tabs**: Email+Folders, Stages+Models, Run Settings, Live Log
- **tooltips** (help=) על כל שדה
- **Progress indicator** (`st.status`) בזמן ריצה
- **Reset confirmation**: דו-שלבי — לחיצה ראשונה מזהירה, שנייה מאשרת
- **Live Log**: @st.fragment(run_every=5), status bar, HTML container

### 📊 Dashboard Page (936 שורות)
- **Period filter**: היום/שבוע/חודש/הכל/טווח מותאם
- **10 KPI cards** בשתי שורות עם delta indicators:
  - Row 1: תקופה → מיילים → שורות → דיוק מיילים → דיוק שורות
  - Row 2: זמן/מייל → זמן/שורה → עלות/מייל → עלות/שורה → עלות כוללת
- **6 tabs**: דיוק, יעילות, לקוחות, שולחים, הודעות אחרונות, ייצוא
- **Weights editor**: expander לעריכת משקלות דיוק (שמירה ל-.env)
- **Human verification**: `st.data_editor` עם עמודת ✓ אימות (שמירה ל-JSONL)
- **Plotly charts**: 14-day accuracy trend, daily cost breakdown, distribution bars
- **Excel export**: 5 sheets (סיכום, מיילים, יומי, לקוחות, שולחים)

### תצורה (.streamlit/config.toml)
```toml
[theme]       # dark, orange primary (#FF8C00)
[server]      # headless, port 8501, 0.0.0.0
[browser]     # gatherUsageStats = false
[client]      # toolbarMode = "minimal"
```

---

## ⚙️ Pipeline — תהליך עיבוד

### Pipeline לכל שרטוט (`extract_drawing_data`):

| # | שלב | מודול | פרומפט |
|---|------|-------|--------|
| 0 | Layout | `stages_generic` | `01_identify_drawing_layout.txt` |
| 0.5 | Rotation | `processing` | `10_detect_rotation.txt` |
| 1 | Basic Info | `stages_generic` | `02_extract_basic_info.txt` |
| 2 | Processes | `stages_generic` | `03_extract_processes.txt` |
| 3 | Notes | `stages_generic` | `04_extract_notes_text.txt` |
| 4 | Area | `stages_generic` | `05_calculate_geometric_area.txt` |
| 5 | Fallback | `stages_rafael` | `06b_extract_processes_from_notes.txt` |
| 9 | Merge | `stage9_merge` | `09_merge_work_descriptions.txt` |

---

## 🔌 תלויות חיצוניות

| חבילה | שימוש |
|--------|-------|
| `openai` | Azure OpenAI Vision API |
| `streamlit` | Web UI framework |
| `plotly` | Dashboard charts |
| `pandas` | DataFrames + Excel export |
| `openpyxl` | Excel styling, BOM catalogs |
| `opencv-python` | Image processing, OCR preprocessing |
| `Pillow` | Image manipulation |
| `pdfplumber` | PDF reading |
| `pytesseract` | OCR (Tesseract) |
| `python-dotenv` | .env loading |
| `msal` | Microsoft Graph auth |

---

## 📝 קבצי הגדרות

| קובץ | תפקיד |
|-------|-------|
| `.env` | Azure OpenAI creds, model config, pricing, ACCURACY_WEIGHT_* |
| `automation_config.json` | הגדרות הרצה אוטומטית + heavy email |
| `automation_state.json` | processed IDs, next run time |
| `email_config.json` | הגדרות חיבור דואר |
| `.streamlit/config.toml` | Streamlit theme, server, client |
| `prompts/*.txt` | 15 פרומפטים AI |
| `BOM/COLORS.xlsx` | קטלוג צבעים |
| `BOM/INSERTS.xlsx` | קטלוג קשיחים |

---

## 📈 היסטוריה

| תאריך | שינוי |
|--------|-------|
| 21/03/2026 | עדכון ספירות שורות, הוספת קבצי backend חסרים, 2 קבצי טסטים חדשים (sanity_checks, drawing_pipeline) |
| 21/03/2026 | רענון תיעוד: עדכון תאריכים ותיאורי אבחון בתיקיית DIAGNOSTICS |
| 25/03/2026 | ניקוי קבצים שלא בשימוש (TEMP/, גיבויים, סקריפטים חד-פעמיים), עדכון כל התיעוד, MANIFEST ו-CHANGELOG |
| 01/03/2026 | ריפקטורינג: חילוץ 6 מודולים (4,018 → 2,059 שורות) |
| 03/2026 | Stage 9, Color/Insert lookups, Insert Validator |
| 07/03/2026 | B2B field 11 → merged_description |
| 09/03/2026 | Heavy Email + Process Analysis + Graph categories |
| 10-12/03/2026 | ★ Streamlit Web UI (3 pages, backend, brand, auto-refresh) |
| 12/03/2026 | 10 KPI + deltas, weights editor, human verification, tooltips, reset confirmation, progress, cost header, toolbarMode=minimal, GPT-4o fallback |
