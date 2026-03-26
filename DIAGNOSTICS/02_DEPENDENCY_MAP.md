# DrawingAI Pro — מפת תלויות בין מודולים

> עדכון אחרון: 25/03/2026 — **כולל Streamlit Web UI layer**

## 📊 תרשים תלויות כללי

```
                    ┌─────────────────────────────────────┐
                    │  customer_extractor_v3_dual.py      │
                    │  (2,160 שורות — Orchestrator)        │
                    │  extract_customer_name() thin wrap  │
                    │  scan_folder() main loop            │
                    └──────────────┬──────────────────────┘
                                   │ imports from ↓
       ┌──────────┬──────────┬─────┼─────┬──────────┬──────────┐
       ▼          ▼          ▼     ▼     ▼          ▼          ▼
  ┌─────────┐┌─────────┐┌────────┐┌───────┐┌──────────┐┌─────────┐
  │extraction││ image/  ││  ai/   ││ file/ ││reporting/││  core/  │
  │ 16 mods ││process. ││vision_││file_u.││excel_exp  ││constants│
  │         ││ 812 ln  ││api    ││classif││pl_gen    ││  90 ln  │
  └─────────┘└─────────┘└────────┘└───────┘└──────────┘└────┬────┘
                                                            │
                    imported by ALL service modules ────────┘
```

## ★ Streamlit Web UI — שכבת תלויות

```
  ┌──────────────────────────────────────────────────────┐
  │  streamlit_app/pages/                                │
  │  ├── 1_🚀_Automation.py ──→ config_manager           │
  │  │                        ──→ runner_bridge           │
  │  │                        ──→ log_reader              │
  │  │                        ──→ email_helpers            │
  │  │                        ──→ brand                   │
  │  ├── 2_📊_Dashboard.py   ──→ log_reader              │
  │  │                        ──→ brand                   │
  │  └── 3_📧_Email.py       ──→ config_manager           │
  │                            ──→ email_helpers            │
  │                            ──→ brand                   │
  └──────────────────────────────────────────────────────┘
          │                         │
  ┌───────▼─────────┐      ┌───────▼────────────┐
  │ backend/        │      │ brand.py            │
  │ config_manager  │──→   │ sidebar_logo()      │
  │   automation_   │      │ brand_header()      │
  │   config.json   │      │ BRAND_CSS           │
  │ runner_bridge   │──→   └─────────────────────┘
  │   Automation    │
  │   Runner.py     │
  │ log_reader      │──→ automation_log.jsonl + status_log.txt
  │ email_helpers   │──→ src.services.email.graph_helper
  └─────────────────┘
```

### Backend Module Dependencies

| Module | Depends On | Used By |
|--------|-----------|---------|
| `config_manager.py` | `automation_config.json` (file I/O) | Automation, Email pages |
| `runner_bridge.py` | `automation_runner.AutomationRunner` | Automation page |
| `log_reader.py` | `automation_log.jsonl`, `status_log.txt`, `logs/*.log` | Dashboard, Automation (header) |
| `email_helpers.py` | `src.services.email.graph_helper.GraphAPIHelper` | Automation, Email pages |
| `brand.py` | PIL (logo), streamlit | All pages |

---

## 🏗️ ארכיטקטורת extraction/ לאחר ריפקטורינג + Stage 9

```
  customer_extractor_v3_dual.py
       │
       ├─→ drawing_pipeline.py ─────→ extract_drawing_data()
       │        ├─→ pn_voting.py
       │        ├─→ sanity_checks.py
       │        ├─→ post_processing.py
       │        ├─→ stages_generic/rafael/iai
       │        ├─→ ocr_engine.py
       │        └─→ image/processing.py
       │
       ├─→ quantity_matcher.py
       ├─→ file_renamer.py
       ├─→ stage9_merge.py ──→ vision_api + color_price_lookup
       ├─→ insert_validator.py
       └─→ insert_price_lookup.py
```

---

## 🔗 תלויות פנימיות — Core

### `src/core/constants.py` (HUB — מיובא ע"י כולם)
```
exports: debug_print, MODEL_RUNTIME, DRAWING_EXTS, STAGE_*, MAX_FILE_SIZE_MB, etc.
← imported by: vision_api, processing, filename_utils, document_reader,
               file_utils, classifier, pl_generator, excel_export,
               stages_generic, stages_rafael, stages_iai, ocr_engine
```

### `src/utils/prompt_loader.py`
```
exports: load_prompt(name: str) → str   (cached @lru_cache)
← imported by: stages_generic, stages_rafael, document_reader,
               classifier, processing, pl_generator
```

### `src/services/ai/vision_api.py`
```
→ src.core.constants, src.services.ai.model_runtime
exports: _build_client, _get_client_for_model, _resolve_stage_call_config,
         _chat_create_with_token_compat, _call_vision_api_with_retry, _calculate_stage_cost
← imported by: classifier, document_reader, pl_generator, processing,
               stages_generic, stages_rafael, stages_iai, ocr_engine, drawing_pipeline
```

### `src/services/ai/model_runtime.py` (254 שורות, גדל מ-143)
```
exports: ModelRuntimeConfig, build_azure_client, calculate_token_cost,
         get_model_endpoint, get_model_api_key, get_model_api_version,
         get_model_deployment, is_model_openai_compat, is_model_reasoning
← imported by: vision_api
```

---

## 🔗 תלויות — Extraction

### `src/services/extraction/drawing_pipeline.py` (844 שורות)
```
→ image/processing, vision_api, ocr_engine
→ stages_generic, stages_rafael, stages_iai
→ pn_voting, sanity_checks, post_processing
← imported by: customer_extractor_v3_dual
```

### `src/services/extraction/filename_utils.py` (594 שורות — HUB שני)
```
← imported by: document_reader, file_utils, excel_export, pl_generator,
               ocr_engine, pn_voting, sanity_checks, quantity_matcher,
               customer_extractor_v3_dual  (9 consumers)
```

### `src/services/extraction/stage9_merge.py`
```
→ vision_api, color_price_lookup (lazy)
← imported by: customer_extractor_v3_dual
```

### `src/services/extraction/color_price_lookup.py`
```
standalone — openpyxl only (BOM/COLORS.xlsx)
← imported by: stage9_merge (lazy)
```

### `src/services/extraction/insert_price_lookup.py`
```
standalone — openpyxl only (BOM/INSERTS.xlsx)
← imported by: customer_extractor_v3_dual
```

---

## 🔗 תלויות — Email

### `src/services/email/*`
```
graph_mailbox.py  ← 916 שורות — Graph API connector + category management
graph_helper.py   ← 528 שורות — Graph utilities + replace_category
graph_auth.py     ← 261 שורות — Graph authentication (MSAL)
shared_mailbox.py ← 593 שורות — EWS connector
factory.py        ← 168 שורות — connector factory
← imported by: automation_runner, email_panel_gui, streamlit email_helpers
```

---

## ⚠️ נקודות לתשומת לב

1. **constants.py**: HUB — כל מודולי השירות תלויים בו
2. **filename_utils.py**: HUB שני — 9 consumers
3. **drawing_pipeline.py**: fan-out גבוה ביותר (7 dependencies)
4. **customer_extractor_v3_dual.py**: orchestrator — מייבא מ-15+ מודולים
5. **stage9_merge.py**: lazy import ל-color_price_lookup (performance)
6. **Streamlit backend**: שכבה דקה מעל core engine — config_manager, runner_bridge, log_reader
7. **runner_bridge.py**: thread-safe wrapper — Streamlit reruns vs. long-running automation
8. **log_reader.py**: שני מקורות נתונים — JSONL (statistics) + status_log.txt (live log)
9. **Dual UI**: Streamlit + Tkinter פועלים במקביל (legacy migration)
