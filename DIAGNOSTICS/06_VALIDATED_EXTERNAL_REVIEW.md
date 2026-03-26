# DrawingAI Pro — Validated External Review

> תאריך: 21/03/2026
> מטרת המסמך: לסנן את דוח האבחון החיצוני לגרסה קצרה שמתבססת רק על טענות שניתן לאמת מול הקוד הנוכחי.

---

## מסקנה קצרה

האבחון החיצוני נראה **הגיוני ברמת הכיוון**, אבל **לא מספיק אמין ברמת המספרים והטענות האבסולוטיות**.

הערכת עבודה:

| היבט | הערכה |
|------|-------|
| כיוון ארכיטקטוני | טוב |
| דיוק מספרי | בינוני-נמוך |
| שימושיות לתיעדוף | טובה |
| התאמה כ-audit פורנזי | חלשה |

---

## 1. מה אומת כנכון

### 1.1 יש ריכוזיות גבוהה בליבת המערכת

הדוח צודק עקרונית בכך שיש מספר orchestrators כבדים מאוד, למשל:

- [customer_extractor_v3_dual.py](customer_extractor_v3_dual.py#L322) — `scan_folder`
- [automation_runner.py](automation_runner.py#L859) — `_run_once_internal`
- [src/services/extraction/drawing_pipeline.py](src/services/extraction/drawing_pipeline.py#L118) — `extract_drawing_data`
- [src/services/extraction/sanity_checks.py](src/services/extraction/sanity_checks.py#L86) — `run_pn_sanity_checks`
- [src/services/extraction/quantity_matcher.py](src/services/extraction/quantity_matcher.py#L58) — `match_quantities_to_drawings`

גם בלי לאמת את מספרי ה-cyclomatic complexity, עצם המבנה מעיד על ריכוזיות גבוהה וסיכון תחזוקתי.

### 1.2 תיעוד DIAGNOSTICS לא מייצג במדויק את המצב הנוכחי

הדוח צדק בכך שחסרים קבצים חדשים ושהמספרים בחלק מהקבצים התיישנו. למשל, [DIAGNOSTICS/01_PROJECT_OVERVIEW.md](DIAGNOSTICS/01_PROJECT_OVERVIEW.md#L37) לא כלל את:

- [streamlit_app/backend/excel_report_builder.py](streamlit_app/backend/excel_report_builder.py)
- [streamlit_app/backend/report_exporter.py](streamlit_app/backend/report_exporter.py)

בנוסף, מספרי שורות שתועדו במסמך היו ישנים ביחס לקבצים בפועל.

### 1.3 התלויות אינן נעולות לפרודקשן

הטענה הזאת נכונה. [requirements.txt](requirements.txt#L5) משתמש ב-`>=` עבור התלויות המרכזיות, ללא lockfile וללא upper bounds.

זו לא שגיאה מיידית, אבל זה כן סיכון תפעולי אמיתי.

### 1.4 קיימת כפילות UI בין Tkinter ל-Streamlit

הדוח צודק שיש שתי שכבות UI חיות במקביל:

- Streamlit: [streamlit_app/app.py](streamlit_app/app.py)
- Tkinter/CustomTkinter: [customer_extractor_gui.py](customer_extractor_gui.py), [automation_main.py](automation_main.py), [automation_panel_gui.py](automation_panel_gui.py), [dashboard_gui.py](dashboard_gui.py), [email_panel_gui.py](email_panel_gui.py)

זה מייצר עלות תחזוקה, גם אם ההגירה עדיין לא הושלמה.

### 1.5 קיימת התפזרות של קריאות env/config

בדיקה ישירה הראתה הרבה קריאות ל-`os.getenv` וקבצים שונים שקוראים `.env`, למשל:

- [src/services/ai/model_runtime.py](src/services/ai/model_runtime.py#L61)
- [src/core/config.py](src/core/config.py#L31)
- [streamlit_app/backend/log_reader.py](streamlit_app/backend/log_reader.py#L116)
- [streamlit_app/backend/excel_report_builder.py](streamlit_app/backend/excel_report_builder.py#L10)
- [streamlit_app/pages/2_📊_Dashboard.py](streamlit_app/pages/2_📊_Dashboard.py#L19)

המספר המדויק שהדוח נתן לא אומת, אבל המסקנה עצמה נכונה: ההגדרות מפוזרות מדי.

---

## 2. מה חלקית נכון

### 2.1 מצב הבדיקות חלש יותר ממה שרצוי, אבל לא "0%" על כל מה שהדוח רמז

יש בפרויקט 25 קבצי בדיקות תחת [tests](tests), ולא מעט בדיקות לוגיקה שימושית, כולל:

- [tests/test_scan_folder_flow.py](tests/test_scan_folder_flow.py#L1)
- [tests/test_heavy_email.py](tests/test_heavy_email.py#L1)
- [tests/test_image_processing.py](tests/test_image_processing.py)
- [tests/test_quantity_matcher.py](tests/test_quantity_matcher.py)
- [tests/test_pl_part_number.py](tests/test_pl_part_number.py#L3)
- [tests/test_structured_bom.py](tests/test_structured_bom.py#L6)

מצד שני, עדיין נראה שיש פער כיסוי אמיתי במודולי ליבה מסוימים. לכן הכיוון של הביקורת נכון, אבל הניסוח בדוח היה חד מדי.

### 2.2 יש בעיות מדידה במספרי השורות, אבל עצם הטענה על drift בתיעוד נכונה

למשל, הדוח צדק בכך שחלק מהקבצים גדלו, אבל כמה מהמספרים שהוא ציין לא תואמים במדויק למצב הנוכחי. דוגמאות מה-workspace הנוכחי:

- [streamlit_app/brand.py](streamlit_app/brand.py) — 492 שורות
- [streamlit_app/backend/runner_bridge.py](streamlit_app/backend/runner_bridge.py) — 306 שורות
- [streamlit_app/pages/1_🚀_Automation.py](streamlit_app/pages/1_🚀_Automation.py) — 715 שורות
- [streamlit_app/pages/2_📊_Dashboard.py](streamlit_app/pages/2_📊_Dashboard.py) — 936 שורות
- [src/services/extraction/stage9_merge.py](src/services/extraction/stage9_merge.py) — 386 שורות

---

## 3. מה לא אומת או נראה שגוי

### 3.1 הטענה "0 bare excepts" שגויה

נמצאו bare excepts לפחות ב:

- [email_panel_gui.py](email_panel_gui.py#L317)

זה לבדו מראה שהדוח לא מספיק מדויק כדי לשמש audit קשיח בלי אימות נוסף.

### 3.2 הציונים והמדדים נראים heuristic

הדוח נותן ציון כולל `62/100` וערכים כמו `73% type hints coverage`, אבל ללא מתודולוגיה שניתנת לשחזור.

לכן נכון להתייחס לציונים כאל אינדיקציה כללית, לא כמדד הנדסי מחייב.

### 3.3 ספירת הקוד הכוללת בדוח לא יציבה

ב-workspace הנוכחי נספרו 113 קבצי Python מחוץ ל-`.venv`, ו-25 קבצי בדיקות תחת [tests](tests). לכן כל מספר כולל בדוח שמוצג כאילו הוא סופי צריך להילקח בזהירות.

---

## 4. פסק דין

אם משתמשים בדוח כדי להבין **איפה לחפש בעיות**, הוא שימושי.

אם משתמשים בו כדי להחליט ש-"המספרים האלו מוכחים", הוא לא מספיק אמין.

פסק דין פרקטי:

1. לקבל את הכיוון האדריכלי שלו.
2. לא לקבל בלי אימות את כל הספירות, האחוזים והציונים.
3. לבנות תוכנית עבודה מתוך הממצאים המאומתים בלבד.

---

## 5. סדר עדיפויות מומלץ

1. לעדכן את מסמכי [DIAGNOSTICS](DIAGNOSTICS) כך שישקפו את מבנה הקוד הנוכחי.
2. ליצור lockfile או strategy מסודרת ל-dependency pinning.
3. להרחיב בדיקות סביב orchestrators וזרימות end-to-end קריטיות.
4. לרכז config/env access לשכבה אחת.
5. להחליט אם Tkinter נשאר זמנית או מתחיל מסלול הוצאה מסודר.

---

*מסמך זה מבוסס על אימות ידני מול ה-workspace הנוכחי, ולא על קבלת הדוח החיצוני כפי שהוא.*