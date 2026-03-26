# דוח Excel אוטומטי — `schedule_report_latest.xlsx`

---

## מה זה?
דוח Excel שנוצר **אוטומטית** בסוף כל סבב אוטומציה.  
מכיל 12 גיליונות: 6 ל"יום נוכחי" + 6 ל"היסטוריה" — מיילים, דיוק, עלויות, זמנים ועוד.

---

## איפה הקובץ?

```
C:\DrawingAI\reports\schedule_report_latest.xlsx
```

הנתיב מוגדר ב-`automation_config.json` במפתח:
```json
"scheduler_report_folder": "C:/DrawingAI/reports"
```

> ⚠️ **חשוב:** ודאי שהנתיב מתאים לשרת ולא למחשב המקומי.  
> לא לכתוב `C:/Users/yelena/Desktop/automation/reports` — זה נתיב מקומי ולא יעבוד בשרת!

---

## איך הוא נוצר?

נוצר אוטומטית — **אין צורך בפעולה ידנית**:

```
Scheduler מריץ סבב (regular / heavy)
        ↓
סבב מסתיים
        ↓
export_schedule_report() נקרא
        ↓
קורא נתונים מ-automation_log.jsonl
        ↓
בונה workbook עם 12 גיליונות
        ↓
שומר ל-schedule_report_latest.xlsx
        ↓
הקובץ הקודם נדרס (תמיד קובץ אחד)
```

---

## דרישות כדי שהדוח ייווצר

- ✅ `scheduler_report_folder` מוגדר בנתיב תקין ב-`automation_config.json`
- ✅ תיקיית היעד קיימת (`C:\DrawingAI\reports\`)
- ✅ `automation_log.jsonl` קיים ומכיל רשומות ריצה
- ✅ Scheduler מופעל (`"scheduler_enabled": true`)

---

## יצירת תיקיית הדוחות

```powershell
New-Item -ItemType Directory -Path "C:\DrawingAI\reports" -Force
```

---

## פתרון בעיות

| בעיה | פתרון |
|------|-------|
| הדוח לא נוצר | בדקי ש-`scheduler_report_folder` ב-`automation_config.json` מצביע לנתיב בשרת |
| תיקייה לא קיימת | `New-Item -ItemType Directory -Path "C:\DrawingAI\reports" -Force` |
| הדוח ריק / חסר נתונים | ודאי ש-`automation_log.jsonl` הועתק מתיקיית `server_copy\` |
| Permission denied | ודאי שלמשתמש שמריץ את המערכת יש הרשאת כתיבה לתיקייה |
| "יום נוכחי" ריק | המערכת מסננת לפי תאריך היום — הדוח יתמלא רק לאחר שיש סבב ביום הנוכחי |
