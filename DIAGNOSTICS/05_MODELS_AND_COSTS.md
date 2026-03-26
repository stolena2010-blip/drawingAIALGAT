# DrawingAI Pro — Models & Costs

> עדכון אחרון: 25/03/2026

---

## 📋 מודלים בשימוש

| Model | Type | Input $/1M tokens | Output $/1M tokens | שימוש |
|-------|------|-------------------:|--------------------:|-------|
| **gpt-5.4** | Vision+Reasoning | $2.50 | $15.00 | שלבים קריטיים (1, 2, 3, 5, 8) |
| **gpt-4o-vision** | Vision | $5.00 | $20.00 | שלבים 0, 4 (תמונות) |
| **gpt-4o** | Vision | $2.50 | $10.00 | Global fallback deployment |
| **gpt-4o-mini** | Text-only | $0.15 | $0.60 | שלבים פשוטים (0, 4, 5, 6) — ב-.env.example |
| **gpt-4o-mini-email** | Text-only | — | — | שלב 6 (Parts List, מיילים) |
| **o4-mini** | Reasoning | $1.10 | $4.40 | שלבים 7, 9 (חשיבה) |
| **gpt-5.2** | Vision+Reasoning | $1.75 | $14.00 | זמין אך לא מוקצה כרגע |

---

## 🔄 שלב ← מודל ← מטרה

### תצורה חיה (automation_config.json)

| Stage | שם | מודל (Live) | מטרה | Max Tokens | Temp |
|:-----:|-----|-------------|-------|:----------:|:----:|
| **0** | Classification / Rotation | `gpt-4o-vision` | סיווג מסמך, זיהוי כיוון סיבוב, מיקום title block | 300 | 0 |
| **1** | Basic Info | **`gpt-5.4`** | חילוץ מק"ט, מספר שרטוט, לקוח, גרסה (שלב קריטי ביותר) | 1,000 | 0 |
| **2** | Processes | **`gpt-5.4`** | חילוץ חומר, ציפוי, צביעה, מפרטים | 1,500 | 0 |
| **3** | Notes | **`gpt-5.4`** | קריאת הערות מהשרטוט | 1,200 | 0 |
| **4** | Area | `gpt-4o-vision` | חישוב שטח גאומטרי של החלק | 800 | 0 |
| **5** | Validation | **`gpt-5.4`** | ולידציה/fallback — חילוץ תהליכים מהערות | 400 | 0 |
| **6** | Parts List | `gpt-4o-mini-email` | קריאת מיילים ומסמכי PL | 2,000 | 0.2 |
| **7** | Email Quantities | `o4-mini` | חילוץ כמויות חכם ממיילים (reasoning) | 800 | 0 |
| **8** | Quote/Order | **`gpt-5.4`** | זיהוי פריטים מהצעות מחיר/הזמנות | 1,500 | 0 |
| **9** | Description Merge | `o4-mini` | מיזוג 5 מקורות תיאור ל-4 שדות מובנים | 32,000 | 0 |

### ברירות מחדל (.env.example) מול תצורה חיה

| Stage | .env.example | Live Config | שינוי |
|:-----:|:-------------|:------------|:------|
| 0 | `gpt-4o-mini` | `gpt-4o-vision` | ⬆ שדרוג |
| 1 | `gpt-5.4` | `gpt-5.4` | — |
| 2 | `gpt-4o` | `gpt-5.4` | ⬆ שדרוג |
| 3 | `gpt-4o` | `gpt-5.4` | ⬆ שדרוג |
| 4 | `gpt-4o-mini` | `gpt-4o-vision` | ⬆ שדרוג |
| 5 | `gpt-4o` | `gpt-5.4` | ⬆ שדרוג |
| 6 | `gpt-4o-mini` | `gpt-4o-mini-email` | ~ התאמה |
| 8 | `gpt-4o` | `gpt-5.4` | ⬆ שדרוג |

---

## 📄 קבצי Prompt ← מיפוי שלבים

| קובץ Prompt | Stage |
|-------------|:-----:|
| `09_classify_document_type.txt` | 0 (סיווג) |
| `10_detect_rotation.txt` | 0 (סיבוב) |
| `01_identify_drawing_layout.txt` | 0.5 (Layout) |
| `02_extract_basic_info.txt` | 1 |
| `06_extract_basic_info_rafael.txt` | 1 (וריאנט רפאל) |
| `03_extract_processes.txt` | 2 |
| `04_extract_notes_text.txt` | 3 |
| `05_calculate_geometric_area.txt` | 4 |
| `06b_extract_processes_from_notes.txt` | 5 (fallback) |
| `11_extract_pl_fields.txt` | 6 |
| `12_analyze_pl_bom.txt` | 6 (BOM) |
| `07_extract_quantities_from_email.txt` | 7 |
| `07b_extract_quantities_fallback.txt` | 7 (fallback) |
| `08_extract_item_details_from_orders.txt` | 8 |
| `09_merge_work_descriptions.txt` | 9 |

---

## 💲 אומדן עלויות — תצורה חיה (3,000 פריטים/חודש)

| Stage | מודל | Tokens ממוצע (in/out) | $/פריט | $/חודש |
|:-----:|-------|:---------------------:|-------:|-------:|
| 0 | gpt-4o-vision | 400 / 35 | $0.0014 | $4.05 |
| **1** | **gpt-5.4** | 3,300 / 230 | $0.0117 | $35.10 |
| 2 | gpt-5.4 | 4,300 / 330 | $0.0141 | $42.15 |
| 3 | gpt-5.4 | 3,650 / 265 | $0.0118 | $35.33 |
| 4 | gpt-4o-vision | 1,660 / 165 | $0.0058 | $17.40 |
| 5 | gpt-5.4 | 330 / 46 | $0.0002 | $0.50 |
| 6 | gpt-4o-mini-email | 465 / 53 | $0.0001 | $0.30 |
| 7 | o4-mini | 530 / 100 | $0.0010 | $3.07 |
| 8 | gpt-5.4 | 330 / 33 | $0.0012 | $3.47 |
| 9 | o4-mini | 1,660 / 1,000 | $0.0062 | $18.68 |
| | | | | |
| **סה"כ** | | **~16,625 / ~2,257** | **$0.053** | **~$160** |

### נוסחת חישוב עלות

$$\text{cost} = \frac{\text{input\_tokens}}{1{,}000{,}000} \times \text{input\_price} + \frac{\text{output\_tokens}}{1{,}000{,}000} \times \text{output\_price}$$

### המרה לש"ח

$$\text{cost\_ils} = \text{cost\_usd} \times 3.7$$

> סה"כ חודשי: **~$160 ≈ ~₪590**

---

## 🔀 שרשרת Fallback

```
  Stage Model (e.g. gpt-5.4)
     │
     ├─ 429 Rate Limit → המתנה retry_after → ניסיון חוזר
     │     └─ עדיין נכשל → fallback ל-gpt-4o (AZURE_DEPLOYMENT)
     │
     ├─ 400 Invalid Prompt → fallback ל-gpt-4o
     │
     ├─ 404 Resource Not Found → fallback ל-gpt-4o
     │
     ├─ Content Filter → ניסיון חוזר עם low-detail + disclaimer
     │
     └─ שגיאה לא מטופלת → fallback ל-gpt-4o
```

### Fallbacks ייחודיים לשלבים

| שלב | Fallback |
|:---:|----------|
| **Stage 1** | `extract_stage1_with_retry` — עד 3 ניסיונות עם אסטרטגיות שונות (DPI גבוה, חיתוך גדול) |
| **Stage 2** | עד 3 ניסיונות; אם הכל נכשל → Stage 5 (חילוץ מהערות במקום מתמונה) |
| **Stage 7** | prompt ראשי → `07b_extract_quantities_fallback.txt` |

---

## ⚙️ מנגנון תמחור (3-tier)

מוגדר ב-`ModelRuntimeConfig` (`src/services/ai/model_runtime.py`):

```
עדיפות 1: STAGE_N_INPUT_PRICE / STAGE_N_OUTPUT_PRICE
           (override ספציפי לשלב)

עדיפות 2: MODEL_<NORMALIZED>_INPUT_PRICE / _OUTPUT_PRICE
           (מחיר לפי שם מודל)

עדיפות 3: AZURE_MODEL_INPUT_PRICE_PER_1M / _OUTPUT_PRICE_PER_1M
           (fallback גלובלי)
```

### CostTracker Class

```python
# src/core/cost_tracker.py
class CostTracker:
    add_usage(input_tokens, output_tokens, cost=None)
    get_summary() → {"total_cost_usd": float, "total_cost_ils": float, ...}
```

- צובר `total_input_tokens`, `total_output_tokens`, `_accumulated_cost`
- כשה-cost מחושב מראש לפי שלב — משתמש ישירות; אחרת מחשב לפי מודל בסיס
- לוג בזמן אמת: `💲 Stage N | model=X | tokens: N in / N out | cost=$X.XXXX`

---

## 📁 קבצי תצורה רלוונטיים

| קובץ | תפקיד |
|-------|--------|
| `.env` (לא ב-Git) | הגדרות מודלים ומחירים בפועל |
| `.env.example` | תבנית לכל משתני הסביבה |
| `automation_config.json` | `stage_models` + `selected_stages` (GUI) |
| `src/services/ai/model_runtime.py` | `ModelRuntimeConfig` — לוגיקת תמחור |
| `src/services/ai/vision_api.py` | קריאות API, חישוב עלות, fallback chains |
| `src/core/cost_tracker.py` | צבירת עלויות בריצה |
| `src/core/constants.py` | קבועי שלבים ושמות תצוגה |

---

*Generated by DrawingAI Pro Diagnostic Tools — refreshed 25/03/2026*
