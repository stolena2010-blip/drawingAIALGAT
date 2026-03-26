# DrawingAI Pro — מדריך התקנה והעלאה לשרת

> עדכון אחרון: 25/03/2026

---

## 📋 דרישות מקדימות בשרת

| רכיב | גרסה מינימלית | הערות |
|-------|---------------|-------|
| Windows Server | 2019+ | נבדק על Windows Server 2022 |
| Python | 3.10+ | מומלץ 3.11 |
| Git | 2.40+ | להתקנה ועדכונים |
| Tesseract OCR | 5.3+ | כולל חבילת שפה `heb` (עברית) |
| NSSM | 2.24+ | אופציונלי — להרצה כ-Windows Service |

### התקנת Tesseract
1. הורד מ: https://github.com/UB-Mannheim/tesseract/wiki
2. בהתקנה — סמן את `Hebrew` בחבילות השפה
3. נתיב ברירת מחדל: `C:\Program Files\Tesseract-OCR\tesseract.exe`

---

## 📁 מה להעתיק לשרת

### אפשרות א' — Clone מ-GitHub (מומלץ)
```powershell
git clone https://github.com/stolena2010-blip/DRAW-ANALIZER.git C:\DrawingAI
cd C:\DrawingAI
```

### אפשרות ב' — העתקה ידנית
העתיקי את **כל** התיקיות והקבצים הבאים:

```
C:\DrawingAI\
├── .env.example              ← תבנית הגדרות (יש לשנות ל-.env)
├── .streamlit\               ← הגדרות Streamlit (theme, port)
├── automation_config.json    ← הגדרות אוטומציה
├── automation_main.py
├── automation_runner.py
├── customer_extractor_gui.py
├── customer_extractor_v3_dual.py
├── dashboard_gui.py
├── email_config.example.json ← תבנית הגדרות מייל
├── email_connector_ews.py
├── email_panel_gui.py
├── main.py
├── process_analysis.py
├── check_next_run.py
├── requirements.txt
├── requirements.lock
├── Run_GUI.bat
├── Run_Statistics.bat
├── Run_Web.bat
├── automation_log.jsonl       ← לוג ריצות — להעתקה אם רוצים לשמר היסטוריה
├── automation_state.json      ← מצב מערכת (processed IDs) — לשמירת רציפות
├── BOM\                      ← קטלוגים (COLORS.xlsx, INSERTS.xlsx)
├── deploy\                   ← סקריפטי התקנה
├── prompts\                  ← 15 קבצי פרומפט AI
├── src\                      ← כל הקוד (ללא __pycache__)
├── streamlit_app\            ← ממשק Web
└── tests\                    ← בדיקות
```

### ❌ מה לא להעתיק
```
.venv/                        ← ייווצר בשרת
__pycache__/                  ← ייווצר אוטומטית
.pytest_cache/
logs/                         ← ייווצר אוטומטית
NEW FILES/                    ← תוצרים — ייווצר אוטומטית
.git/                         ← רק אם משתמשים ב-clone
status_log.txt                 ← ייווצר אוטומטית
.env                           ← מכיל מפתחות! ייצור ידנית בשרת
```

---

## 🔧 שלבי התקנה

### שלב 1 — התקנה אוטומטית (מומלץ)

פתחי **PowerShell כ-Administrator** בשרת:

```powershell
cd C:\DrawingAI\deploy
.\install_server.ps1 -GitHubRepo "https://github.com/stolena2010-blip/DRAW-ANALIZER.git"
```

הסקריפט:
1. בודק Git ו-Python
2. עושה Clone
3. יוצר venv + מתקין תלויות
4. בודק Tesseract
5. יוצר קבצי קונפיגורציה
6. יוצר תיקיות עבודה
7. רושם כ-Windows Service (אם NSSM קיים)

### שלב 2 — התקנה ידנית (אם אין Clone)

```powershell
# 1. צרי venv
cd C:\DrawingAI
python -m venv .venv

# 2. הפעלת venv
.\.venv\Scripts\Activate.ps1

# 3. התקנת תלויות
pip install -r requirements.txt

# 4. צרי קובץ .env מהתבנית
Copy-Item .env.example .env
# ערכי אותו עם המפתחות שלך!

# 5. צרי email_config.json
Copy-Item email_config.example.json email_config.json

# 6. צרי תיקיות עבודה
New-Item -ItemType Directory -Path logs -Force
```

### שלב 3 — הגדרת קונפיגורציה

#### `.env` — חובה לערוך!
```ini
# מפתחות Azure OpenAI
AZURE_OPENAI_ENDPOINT=https://your-endpoint.openai.azure.com/
AZURE_OPENAI_API_KEY=your-api-key

# Microsoft Graph API (למייל)
GRAPH_TENANT_ID=your-tenant-id
GRAPH_CLIENT_ID=your-client-id
GRAPH_CLIENT_SECRET=your-client-secret

# Tesseract
TESSERACT_PATH=C:/Program Files/Tesseract-OCR/tesseract.exe
```

#### `automation_config.json` — הגדרות אוטומציה
```json
{
  "shared_mailboxes": ["quotes@yourcompany.com"],
  "folder_name": "תיבת דואר נכנס",
  "recipient_email": "target@yourcompany.com",
  "download_root": "C:/DrawingAI/downloads",
  "tosend_folder": "C:/DrawingAI/to_send",
  "output_copy_folder": "C:/DrawingAI/archive",
  "scheduler_report_folder": "C:/DrawingAI/reports",
  "scheduler_enabled": true,
  "poll_interval_minutes": 5,
  "auto_send": true
}
```

### שלב 4 — בדיקה

```powershell
# בדיקת הרצה בסיסית
.\.venv\Scripts\python.exe main.py

# בדיקת Streamlit Web UI
.\.venv\Scripts\python.exe -m streamlit run streamlit_app/app.py --server.port 8501

# הרצת בדיקות
.\.venv\Scripts\python.exe -m pytest tests/ -v
```

---

## 🖥️ הרצה

### אפשרות א' — Streamlit Web UI (מומלץ)
```powershell
Run_Web.bat                   # אוטומטית — פותח דפדפן ב-http://localhost:8501
```

### אפשרות ב' — Windows Service (להרצת רקע)
```powershell
# רישום כשירות (פעם ראשונה)
.\deploy\register_service.ps1 -Action install

# הפעלה
.\deploy\register_service.ps1 -Action start

# עצירה
.\deploy\register_service.ps1 -Action stop

# סטטוס
.\deploy\register_service.ps1 -Action status
```

### אפשרות ג' — GUI ישן (Tkinter)
```powershell
Run_GUI.bat
```

---

## 🔄 שדרוג גרסה

### שדרוג מהיר (דאבל-קליק)
לחצי על `deploy\UPDATE.bat` — מריץ את כל התהליך אוטומטית.

### שדרוג ידני מ-PowerShell
```powershell
cd C:\DrawingAI\deploy
.\update.ps1
```

### מה הסקריפט עושה:
1. **עוצר** את השירות (אם רץ)
2. **מגבה** קבצי קונפיגורציה (`.env`, `automation_config.json`, `automation_state.json`)
3. **מושך** קוד חדש מ-GitHub (`git pull`)
4. **מעדכן** תלויות Python (רק אם `requirements.txt` השתנה)
5. **מפעיל** מחדש את השירות

### אפשרויות נוספות:
```powershell
# שדרוג מ-branch אחר
.\update.ps1 -Branch feature/new-feature

# שדרוג בלי הפעלה מחדש
.\update.ps1 -SkipRestart

# כפיית עדכון תלויות
.\update.ps1 -ForceRequirements
```

### שדרוג ללא Git (העתקה ידנית)
1. **גבי** קבצי קונפיגורציה:
   ```powershell
   Copy-Item .env .env.bak
   Copy-Item automation_config.json automation_config.json.bak
   Copy-Item automation_state.json automation_state.json.bak
   ```
2. **העתיקי** את הקבצים החדשים (ראי רשימה למעלה)
3. **שחזרי** קונפיגורציה:
   ```powershell
   Copy-Item .env.bak .env
   Copy-Item automation_config.json.bak automation_config.json
   Copy-Item automation_state.json.bak automation_state.json
   ```
4. **עדכני** תלויות:
   ```powershell
   .\.venv\Scripts\pip.exe install -r requirements.txt
   ```
5. **הפעילי** מחדש

---

## 📂 מבנה תיקיות בשרת (לאחר התקנה)

```
C:\DrawingAI\                        ← קוד הפרויקט
├── .env                             ← מפתחות (אל תעלי ל-Git!)
├── .venv\                           ← Python virtual environment
├── automation_config.json           ← הגדרות אוטומציה
├── automation_state.json            ← מצב מערכת (נוצר אוטומטית)
├── automation_log.jsonl             ← לוג ריצות (נוצר אוטומטית)
├── logs\                            ← לוגים
├── NEW FILES\                       ← פלט B2B (נוצר אוטומטית)
├── deploy\backups\                  ← גיבויי קונפיגורציה (מ-update.ps1)
└── ...

C:\DrawingAI\downloads\              ← הורדות מייל (מוגדר ב-config)
C:\DrawingAI\to_send\                ← קבצי TO_SEND
C:\DrawingAI\archive\                ← ארכיון
C:\DrawingAI\reports\                ← דוחות Excel (schedule_report_latest.xlsx)
```

---

## ⚠️ הערות חשובות

1. **אל תעלי `.env` ל-Git** — מכיל מפתחות API
2. **לא לדרוס `automation_state.json` בשדרוג** — מכיל processed IDs שלא צריך לאפס (סקריפט `update.ps1` מגבה אוטומטית)
3. **גיבוי אוטומטי** — `update.ps1` מגבה קונפיגורציה לפני כל עדכון ב-`deploy\backups\`
4. **Tesseract** — חייב להיות מותקן בנפרד (לא חלק מ-pip)
5. **Firewall** — אם רוצים גישה ל-Streamlit ממחשבים אחרים, פתחי port 8501
6. **Streamlit בענן** — הגדרות ב-`.streamlit/config.toml` כבר מוגדרות ל-`0.0.0.0` (כל הממשקים)

---

## 🔍 פתרון בעיות

| בעיה | פתרון |
|------|-------|
| `ModuleNotFoundError` | `pip install -r requirements.txt` |
| Tesseract not found | בדקי `TESSERACT_PATH` ב-`.env` |
| Port 8501 תפוס | `netstat -aon \| findstr :8501` → `taskkill /PID <PID> /F` |
| Service לא עולה | בדקי `logs\service_stderr.log` |
| Graph API שגיאה | בדקי `GRAPH_*` ב-`.env`, וודאי הרשאות ב-Azure AD |
| קבצי BOM חסרים | וודאי ש-`BOM\COLORS.xlsx` ו-`BOM\INSERTS.xlsx` קיימים |
