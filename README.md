# DrawingAI Pro

**Automated engineering drawing analysis powered by Azure OpenAI Vision API.**

## Overview

DrawingAI Pro automates the extraction of technical data from engineering drawings received via email. It processes PDF/image files through multi-stage AI analysis, classifies documents, extracts part numbers, materials, processes, and surface treatments, then generates structured Excel and B2B reports.

## Key Features

- **Multi-stage extraction**: basic info, processes, notes, geometric area, merge, validation
- **Customer-specific logic** (IAI / RAFAEL / Generic variants)
- **File classification** (drawings vs documents)
- **B2B export** with confidence levels (LOW/MEDIUM/HIGH)
- **Automated email workflow**: receive → download → analyze → export → send
- **Multi-mailbox support** via Microsoft Graph API
- **Cost tracking** per Azure API stage
- **OCR** with Tesseract + advanced image preprocessing
- **Streamlit Web UI** — real-time dashboard, automation panel, email management
- **Scheduler report** — Excel report generated per automation cycle

## Architecture

```
streamlit_app/                   — ★ Streamlit Web UI (Automation, Dashboard, Email)
├── backend/                     — config_manager, runner_bridge, log_reader, report_exporter
├── pages/                       — 3 Streamlit pages
└── brand.py                     — CSS, logos, RTL support

customer_extractor_v3_dual.py    — Main orchestrator (scan_folder + extraction pipeline)
automation_runner.py             — Automated email processing cycle + heavy email

src/
├── core/                        — Config, constants, cost tracker, exceptions
├── services/
│   ├── ai/vision_api.py         — Azure OpenAI API calls with retry logic
│   ├── image/processing.py      — Image preprocessing, rotation, OCR
│   ├── extraction/              — 16 modules: pipeline stages, OCR, P.N. voting, sanity checks
│   ├── file/                    — File classification, renaming, metadata, TO_SEND operations
│   ├── reporting/               — Excel reports, B2B export, PL generation
│   └── email/                   — Microsoft Graph API email integration
├── models/                      — Drawing models, enums
└── utils/                       — Logger, prompt loader
```

## Quick Start

### Prerequisites

- Python 3.10+
- Tesseract OCR installed
- Azure OpenAI API access

### Installation

```bash
git clone <repo-url>
cd "AI DRAW_STEAMLIT"
pip install -r requirements.txt
cp .env.example .env
# Edit .env with your Azure OpenAI credentials
```

### Running

```bash
Run_Web.bat                 # Streamlit Web UI (recommended)
Run_GUI.bat                 # Legacy Tkinter GUI (Windows)
Run_Statistics.bat          # Process analysis statistics
python main.py              # CLI mode
```

### Running Tests

```bash
pytest
```

## Project Stats

| Metric | Value |
|--------|-------|
| Python files | 102 |
| Prompt templates | 15 |
| Test files | 25 |
| Streamlit pages | 3 |

```bash
python -m pytest tests/ -v
```

## Configuration

All configuration is via `.env` file. See `.env.example` for available options.

## Project Stats

- 64 automated tests
- 10 focused modules
- Multi-mailbox email automation
- 99.7% success rate in production (594 emails processed)
