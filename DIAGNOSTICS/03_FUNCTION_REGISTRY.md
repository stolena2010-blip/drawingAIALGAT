# DrawingAI Pro — Function Registry

> עדכון אחרון: 25/03/2026 — **כולל Streamlit Web UI functions**
> Total files: **111 Python** · Total lines: **~31,600**

---

## 1. Core & Utilities

### src/core/constants.py (90 lines)
| Line | Name | Description |
|------|------|-------------|
| 28 | `debug_print` | Debug message when AI_DRAW_DEBUG enabled |
| — | Constants | DEBUG_ENABLED, MODEL_RUNTIME, DRAWING_EXTS, STAGE_*, MAX_FILE_SIZE_MB |

### src/core/config.py (181 lines)
Configuration classes for project settings.

### src/core/cost_tracker.py (78 lines)
Cost tracking utilities (CostTracker class).

### src/core/exceptions.py (65 lines)
10 custom exception classes.

### src/utils/prompt_loader.py (39 lines)
| Line | Name | Description |
|------|------|-------------|
| 20 | `load_prompt` | Load prompt from `prompts/<name>.txt` (LRU-cached) |

### src/utils/logger.py (141 lines)
| Line | Name | Description |
|------|------|-------------|
| 12 | `ColoredFormatter` | ANSI color formatter |
| 31 | `GUICallbackHandler` | Forward logs to GUI callback |
| 43 | `setup_logging` | Root logging: console, rotating file, GUI |
| 117 | `get_logger` | Get named Logger |
| 130 | `create_log_file` | Create log file path |

---

## 2. AI Services

### src/services/ai/model_runtime.py (254 lines)
| Line | Name | Description |
|------|------|-------------|
| 13 | `_normalize_azure_endpoint` | Strip trailing `/openai/v1` |
| 22 | `_safe_float` | Parse float with fallback |
| 29 | `_safe_int` | Parse int with fallback |
| 36 | `_try_float` | Try float parse → None |
| 43 | `_normalize_model_env_key` | Model name → env-var key |
| 50 | `ModelRuntimeConfig` | Frozen dataclass for Azure config |
| 59 | `  from_env` | Create from env vars |
| 69 | `  get_stage_model` | Model for stage |
| 94 | `  get_stage_input_price` | Input price for stage |
| 103 | `  get_stage_output_price` | Output price for stage |
| 112 | `  get_stage_temperature` | Temperature for stage |
| 118 | `  get_stage_max_tokens` | max_tokens for stage |
| 125 | `build_azure_client` | Build AzureOpenAI client |
| 135 | `calculate_token_cost` | USD cost from tokens |
| — | `get_model_endpoint` | Endpoint for specific model |
| — | `get_model_api_key` | API key for specific model |
| — | `get_model_api_version` | API version for model |
| — | `get_model_deployment` | Deployment name |
| — | `is_model_openai_compat` | OpenAI vs Azure compat check |
| — | `is_model_reasoning` | Is reasoning model (o4-mini etc.) |

### src/services/ai/vision_api.py (279 lines)
| Line | Name | Description |
|------|------|-------------|
| 29 | `_build_client` | Init Azure OpenAI client |
| 34 | `_get_client_for_model` | Get client matching model (GPT-4o fallback) |
| 38 | `_resolve_stage_call_config` | Model/tokens/temp for stage |
| 65 | `_calculate_stage_cost` | USD cost for stage |
| 71 | `_log_stage_completion` | Print per-stage stats |
| 93 | `_chat_create_with_token_compat` | Completion with auto-retry |
| 153 | `_call_vision_api_with_retry` | Vision API + content-filter retry |

---

## 3. Image Processing

### src/services/image/processing.py (812 lines)
| Line | Name | Description |
|------|------|-------------|
| 42 | `_downsample_high_res_image` | Downsample > MAX_IMAGE_DIMENSION |
| 75 | `_enhance_contrast_for_title_block` | Title block contrast |
| 147 | `_extract_image_smart` | Smart extraction: overview + crop |
| 297 | `_assess_image_quality` | Sharpness, contrast, brightness |
| 362 | `_validate_rotation_improvement` | OCR clarity after rotation |
| 439 | `_apply_rotation_angle` | Apply rotation |
| 458 | `_estimate_quarter_turn_hint` | 90°/270° heuristic |
| 491 | `_fix_image_rotation` | Auto-detect and correct rotation |

---

## 4. Extraction Stages

### src/services/extraction/stages_generic.py (437 lines)
| Name | Description |
|------|-------------|
| `identify_drawing_layout` | Stage 0: Layout detection |
| `extract_basic_info` | Stage 1: Customer, P.N., revision |
| `extract_processes_info` | Stage 2: Material, coating, BOM |
| `validate_notes_before_stage5` | Pre-Stage 5 validation |
| `extract_notes_text` | Stage 3: Technical notes |
| `calculate_geometric_area` | Stage 4: Geometric area |

### src/services/extraction/stages_rafael.py (314 lines)
`*_rafael()` variants + `extract_processes_from_notes()` (Stage 5)

### src/services/extraction/stages_iai.py (310 lines)
`_extract_iai_top_red_identifier()` + `*_iai()` variants

### src/services/extraction/stage9_merge.py (373 lines)
| Name | Description |
|------|-------------|
| `merge_descriptions` | Stage 9: מיזוג 5 מקורות → 4 שדות (o4-mini) |
| `_sum_pl_primary_qty` | PL quantity sum |
| `_sum_drawing_primary_qty` | Drawing quantity sum |
| `_calc_hardware_count` | Hardware count |
| `_process_batch` | Process batch of drawings |

---

## 5. Pipeline & Voting & Sanity

### src/services/extraction/drawing_pipeline.py (844 lines)
| Name | Description |
|------|-------------|
| `extract_drawing_data` | Full pipeline: Stages 0-4 + voting + sanity |
| `_run_with_timeout` | Timeout wrapper |

### src/services/extraction/pn_voting.py (238 lines)
| Name | Description |
|------|-------------|
| `deduplicate_line` | Deduplicate P.N. lines |
| `extract_pn_dn_from_text` | Extract P.N./D.N. from OCR text |
| `vote_best_pn` | 3-way voting for best P.N. |

### src/services/extraction/sanity_checks.py (423 lines)
| Name | Description |
|------|-------------|
| `_find_near_match_in_filename` | Fuzzy filename matching |
| `is_cage_code` | CAGE code validation |
| `run_pn_sanity_checks` | Sanity checks A-D |
| `calculate_confidence` | Confidence level: full/high/medium/low/none |

### src/services/extraction/post_processing.py (135 lines)
| Name | Description |
|------|-------------|
| `post_process_summary_from_notes` | Extract processes from notes |

### src/services/extraction/quantity_matcher.py (399 lines)
| Name | Description |
|------|-------------|
| `_key_matches_any_drawing` | Drawing key matching |
| `match_quantities_to_drawings` | Match quantities to drawings |
| `extract_base_and_suffix` | Base P.N. + suffix separation |
| `override_pn_from_email` | Override P.N. from email subject |

---

## 6. OCR & Documents

### src/services/extraction/ocr_engine.py (554 lines)
`MultiOCREngine` class + `extract_stage1_with_retry`

### src/services/extraction/document_reader.py (783 lines)
`_read_email_content`, `_extract_quantities_from_order_pdf`, `_extract_text_via_ocr`, `_extract_item_details_from_documents`

### src/services/extraction/filename_utils.py (594 lines)
13 functions: `check_value_in_filename`, `extract_part_number_from_filename`, `_normalize_item_number`, etc.

---

## 7. File & Pricing Services

### src/services/file/classifier.py (339 lines)
`classify_file_type`

### src/services/file/file_utils.py (708 lines)
9 functions: `_get_file_metadata`, `_detect_text_heavy_pdf`, `_build_drawing_part_map`, etc.

### src/services/file/file_renamer.py (89 lines)
`rename_files_by_classification`

### src/services/extraction/color_price_lookup.py (324 lines)
`has_paint_process`, `lookup_color_prices` + 8 internal helpers

### src/services/extraction/insert_price_lookup.py (172 lines)
`lookup_insert_price`, `enrich_inserts_with_prices`

### src/services/extraction/insert_validator.py (133 lines)
`validate_inserts_hardware`

---

## 8. Reporting

### src/services/reporting/b2b_export.py (258 lines)
`_save_text_summary`, `_save_text_summary_with_variants` (Field 11 = merged_description)

### src/services/reporting/excel_export.py (610 lines)
`_save_classification_report`, `_update_pl_sheet_with_associated_items`, `_save_results_to_excel`

### src/services/reporting/pl_generator.py (942 lines)
`extract_pl_data`, `_generate_pl_summary_hebrew`, `_generate_pl_summary_english`, `_determine_pl_main_part_number`

---

## 9. Email Services

### src/services/email/graph_mailbox.py (916 lines)
26 methods including `ensure_category`, `replace_message_category`

### src/services/email/graph_helper.py (528 lines)
14 methods including `replace_category`

### src/services/email/graph_auth.py (261 lines)
8 methods — MSAL authentication

### src/services/email/shared_mailbox.py (593 lines)
15 methods — EWS connector

### src/services/email/factory.py (168 lines)
`EmailMethod` enum, `EmailConnectorFactory`

---

## 10. ★ Streamlit Web UI

### streamlit_app/brand.py (230 lines)
| Name | Description |
|------|-------------|
| `_logo_b64()` | Load + base64 encode logo |
| `sidebar_logo()` | Sidebar logo + version |
| `brand_header(title, subtitle)` | Orange-green gradient header |
| `BRAND_CSS` | RTL, dark theme, nav styling |

### streamlit_app/backend/config_manager.py (88 lines)
| Name | Description |
|------|-------------|
| `load_config()` | Read automation_config.json |
| `save_config(cfg)` | Write automation_config.json |
| `load_state()` | Read automation_state.json |
| `reset_state()` | Reset state to defaults |
| `_default_config()` | Default config dictionary |

### streamlit_app/backend/email_helpers.py (133 lines)
| Name | Description |
|------|-------------|
| `test_mailbox_connection(mailbox)` | Test single mailbox |
| `test_all_mailboxes(mailboxes)` | Test all mailboxes |
| `load_folders_for_mailbox(mailbox)` | Load Outlook folders |
| `format_folder_label(path, count)` | Format folder display |
| `parse_mailboxes_text(raw_text)` | Parse comma/newline mailboxes |

### streamlit_app/backend/log_reader.py (373 lines)
| Name | Description |
|------|-------------|
| `_read_jsonl_file(path)` | Parse single JSONL file |
| `load_log_entries(max_entries)` | Load + dedupe all JSONL files |
| `filter_by_period(entries, period, ...)` | Filter by date range |
| `get_accuracy_weights()` | Load weights from env |
| `calc_weighted_accuracy(data, weights)` | Weighted accuracy % |
| `save_entry_field(entry_id, field, value)` | ★ Update single field in JSONL |
| `get_latest_log_file()` | Find most recent log file |
| `_parse_log_file_line(line)` | Parse log line format |
| `read_log_tail(n_lines)` | Read status_log.txt + fallback to log files |
| `detect_active_run()` | Detect active run from log timestamps |
| `get_countdown()` | Next run countdown timer |

### streamlit_app/backend/runner_bridge.py (203 lines)
| Name | Description |
|------|-------------|
| `RunnerBridge` | Thread-safe wrapper for AutomationRunner |
| `  start()` | Start continuous loop |
| `  stop()` | Stop loop |
| `  run_once()` | Single regular run (threaded) |
| `  run_heavy()` | Single heavy run (threaded) |
| `  get_run_status()` | Parse status from log lines |
| `  is_busy` | True during one-shot run |

### streamlit_app/pages/1_🚀_Automation.py (658 lines)
| Name | Description |
|------|-------------|
| `_header_status_fragment()` | Auto-refresh header: status + cost |
| `_live_log_fragment()` | Auto-refresh log viewer |
| `_gather_config()` | Collect all form values → dict |

### streamlit_app/pages/2_📊_Dashboard.py (873 lines)
| Name | Description |
|------|-------------|
| `_entry_ts(e)` | Extract timestamp from entry |
| `_acc_color(val)` | Color class for accuracy value |
| `_global_accuracy(entries)` | Row-level weighted accuracy |
| `_email_accuracy(entries)` | Per-email average accuracy |
| `_total_items/cost/time/files(entries)` | Aggregation helpers |
| `_confidence_totals(entries)` | Sum confidence levels |
| `_entries_by_day(entries)` | Group by date |
| `_prev_period_entries(all_ents, period)` | Previous period for deltas |
| `_delta(current, previous)` | Calculate delta value |
| `_delta_str(val)` | Format delta string |
| `_delta_str4(current, previous)` | Format cost delta (4 decimals) |

---

## 11. Legacy GUI (Tkinter)

### automation_panel_gui.py (~1,250 lines)
Full Tkinter automation panel.

### customer_extractor_gui.py (1,303 lines)
Main Tkinter GUI — manual extraction.

### dashboard_gui.py (~1,500 lines)
Tkinter dashboard with statistics.

### email_panel_gui.py (~485 lines)
Email management Tkinter panel.

---

## 12. Tests (25 files)

`test_automation_utils`, `test_b2b_export`, `test_classifier_guards`, `test_color_price_lookup`,
`test_constants`, `test_drawing_pipeline`, `test_excel_export`, `test_file_classification`, `test_file_utils`,
`test_filename_utils`, `test_graph_helper`, `test_heavy_email`, `test_image_processing`,
`test_insert_price_lookup`, `test_insert_validator`, `test_normalize_item`,
`test_part_number_fallback`, `test_pl_part_number`, `test_pn_voting`,
`test_quantity_matcher`, `test_sanity_checks`, `test_scan_folder_flow`, `test_structured_bom`,
`test_validate_notes`, `test_vision_api`
