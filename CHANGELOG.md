# Changelog

כל השינויים המשמעותיים בפרויקט מתועדים כאן.

## [3.2.0] - 2026-03-25

### חדש
- **Streamlit Web UI** — ממשק Web מלא (3 עמודים: Automation, Dashboard, Email)
- **Dashboard** — 10 KPI cards עם delta, 6 tabs, Plotly charts, human verification
- **Automation Page** — 4 tabs, live log, progress indicator, reset confirmation
- **Scheduler Report** — דוח Excel אוטומטי (schedule_report_latest.xlsx) בסוף כל סבב
- **Report Exporter** — backend module לייצוא דוחות מתוזמנים
- **Runner Bridge** — thread-safe wrapper ל-AutomationRunner בסביבת Streamlit
- **Brand Module** — CSS, לוגו, RTL, dark theme (#FF8C00)
- **Weights Editor** — עריכת משקלות דיוק מהדשבורד (שמירה ל-.env)
- **Stage 9 Merge** — מיזוג תיאורים חכם באמצעות o4-mini
- **Color/Insert Price Lookup** — חיפוש מחירי צבע וקשיחים מקטלוג BOM
- **Insert Validator** — אימות קשיחים
- **Quantity Matcher** — חילוץ כמויות ממיילים ומזמינות רכש
- **P.N. Voting** — pdfplumber + Tesseract + Vision voting
- **Sanity Checks** — בדיקות תקינות מתקדמות
- **deploy/** — סקריפטי התקנה (install_server.ps1, register_service.ps1, UPDATE.bat)

### שיפורים
- הוספת 20 קבצי בדיקות (סה"כ 25)
- requirements.lock לנעילת תלויות
- ניקוי קבצים שלא בשימוש (TEMP/, גיבויים, סקריפטים חד-פעמיים)
- עדכון תיעוד ו-MANIFEST

## [3.1.0] - 2026-02-21

### תיקונים
- **PL Detection** — regex לא תפס `PL_TL-4341` (underscore = word char). תיקון: lookahead/lookbehind
- **Text-Heavy Threshold** — סף 2000 מילים חסם שרטוטים עם routing chart. שונה ל-700 מילים/עמוד
- **Text-Heavy V3** — keyword bypass (DRAWING NO, P.N., SCALE → skip check)
- **PO DPI** — 400 DPI × zoom ×3 = 135MP = 3.5 שעות. הופחת ל-200 DPI × zoom ×2 = 11MP = 30 שניות
- **Smart DPI** — שרטוטים ענקיים (>12K px) לא עוברים upscale ל-400 DPI
- **Stop Button** — `_stop_event` לא נבדק בלולאת המיילים. הוספת check בתחילת כל מייל
- **State Save** — באג: per-message save כתב dict ריק. תוקן + logging + fallback
- **Exception Logging** — outer try/except שינה מ-debug ל-error + full traceback

### חדש
- **Dashboard: Items per Mail** — סקציה חדשה עם ממוצע, חציון, התפלגות
- **Dashboard: Reset with Time** — שדה שעה:דקה באיפוס סטטיסטיקה
- **Smart P.N. Voting** — pdfplumber + Tesseract + Vision voting

## [3.0.0] - 2026-02-01

### חדש
- Automation Runner — ניטור אוטומטי של Shared Mailbox
- 5 GUI Panels — Automation, Extractor, Dashboard, Email, Send
- Parts List integration — PL parsing + association
- 3 customer models — Rafael, IAI, Generic
- Confidence system — FULL/HIGH/MEDIUM/LOW
- B2B variants — 3 versions filtered by confidence

## [2.0.0] - 2025-12-01

### חדש
- Multi-stage pipeline (9 stages)
- Azure OpenAI GPT-4o integration
- Microsoft Graph API (replacing EWS)
- OCR engine with Tesseract + pdfplumber
- Disambiguation engine (O↔0, B↔8)
