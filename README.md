# PDF Translator App

A medical document translation tool built for **GlobeMed**. Upload a scanned or digital PDF (Arabic/English mixed), and receive a clean, structured English PDF with all fields extracted, translated, and formatted.

---

## Table of Contents

- [Features](#features)
- [Tech Stack](#tech-stack)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Configuration](#configuration)
- [Running the App](#running-the-app)
- [API Reference](#api-reference)
- [Design Decisions](#design-decisions)
- [Production Checklist](#production-checklist)

---

## Features

- Upload scanned or digital PDFs — Arabic, English, or mixed
- Claude reads PDFs natively — no separate OCR step
- Extracts all visible fields, tables, and handwritten notes
- Translates labels and values to English while preserving numbers, dates, names, and amounts exactly
- Marks uncertain fields with a confidence level (`high` / `medium` / `low`) instead of guessing
- Generates a clean, formatted English PDF
- Exports structured JSON for downstream processing
- Primary + fallback model chain — automatic retry with a more capable model on failure
- Per-job file storage — every translation is isolated by UUID

---

## Tech Stack

| Technology | Version | Role |
|---|---|---|
| **Python** | 3.10+ | Language |
| **FastAPI** | 0.116 | Backend REST API |
| **Uvicorn** | 0.35 | ASGI server |
| **Streamlit** | 1.49 | Frontend UI |
| **Anthropic Claude API** | — | AI extraction + translation |
| **httpx** | 0.28 | HTTP client for Claude API calls |
| **PyMuPDF** | 1.26 | PDF validation and page counting |
| **ReportLab** | 4.4 | Programmatic PDF generation |
| **Pydantic v2** | 2.11 | Data validation and schemas |
| **pydantic-settings** | 2.10 | Type-safe `.env` config |
| **python-dotenv** | 1.1 | `.env` file loading |
| **python-multipart** | 0.0.20 | Multipart file upload support |
| **uv** | latest | Package manager and virtual environment |

---

## Architecture

```
┌─────────────────────────────────────┐
│           Browser / User            │
└────────────────┬────────────────────┘
                 │  http://localhost:8501
                 ▼
┌─────────────────────────────────────┐
│         Streamlit Frontend          │
│  frontend/streamlit_app.py          │
│                                     │
│  - PDF file uploader                │
│  - POSTs to FastAPI backend         │
│  - Renders sections, fields,        │
│    tables, warnings, metrics        │
│  - Download links (PDF + JSON)      │
└────────────────┬────────────────────┘
                 │  POST /api/v1/translate
                 │  multipart/form-data
                 ▼
┌─────────────────────────────────────┐
│         FastAPI Backend             │
│  backend/app/                       │
│                                     │
│  DocumentPipeline.process()         │
│    │                                │
│    ├─ Validate size & page count    │  PyMuPDF
│    ├─ Save original PDF             │  file_store.py
│    ├─ Build system + user prompts   │  extraction_prompt.py
│    ├─ Call Claude (primary model)   │  anthropic_client.py
│    │   └─ Fallback to Opus if fails │
│    ├─ Validate JSON → Pydantic      │  schemas.py
│    ├─ Save translated_document.json │  file_store.py
│    └─ Render translated_document.pdf│  pdf_renderer.py
│                                     │
│  Returns TranslateResponse JSON     │
└────────────────┬────────────────────┘
                 │  HTTPS
                 ▼
┌─────────────────────────────────────┐
│         Anthropic Claude API        │
│                                     │
│  - Receives base64-encoded PDF      │
│  - Reads pages natively (vision)    │
│  - Returns structured JSON          │
└─────────────────────────────────────┘
```

---

## Project Structure

```
pdf_translator_app/
│
├── backend/
│   └── app/
│       ├── api/
│       │   └── routes.py             # HTTP endpoints
│       ├── models/
│       │   └── schemas.py            # Pydantic data models
│       ├── services/
│       │   ├── anthropic_client.py   # Claude API client
│       │   ├── extraction_prompt.py  # System + user prompt builders
│       │   ├── pipeline.py           # Job orchestration
│       │   └── pdf_renderer.py       # ReportLab PDF generation
│       ├── utils/
│       │   ├── file_store.py         # Filesystem I/O per job
│       │   └── json_utils.py         # JSON extraction from raw LLM text
│       ├── config.py                 # Settings (reads from .env)
│       └── main.py                   # FastAPI app + CORS setup
│
├── frontend/
│   └── streamlit_app.py              # Entire UI in one file
│
├── storage/                          # Auto-created; one folder per job UUID
│   └── <job-uuid>/
│       ├── <original-filename>.pdf
│       ├── translated_document.pdf
│       └── translated_document.json
│
├── .env                              # Your secrets — never commit this
├── .env.example                      # Template for required variables
├── .gitignore
├── me.md                             # Full codebase reference
├── requirements.txt                  # Pip-compatible dependency list
├── pyproject.toml                    # Project metadata
└── uv.lock                           # Locked dependency versions
```

---

## Installation

This project uses [**uv**](https://github.com/astral-sh/uv) — a fast Python package manager that handles virtual environments and dependency locking in one tool.

### 1. Install uv

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Or via pip
pip install uv
```

Verify the install:

```bash
uv --version
```

---

### 2. Clone the repository

```bash
git clone <repo-url>
cd pdf_translator_app
```

---

### 3. Create a virtual environment and install dependencies

```bash
uv venv
```

This creates a `.venv/` folder using the Python version specified in `.python-version`.

Then install all dependencies from the lockfile:

```bash
uv sync
```

`uv sync` reads `uv.lock` to install the exact pinned versions — ensuring every machine runs the same environment. This is equivalent to `pip install -r requirements.txt` but faster and reproducible.

> **Why uv?**
> - 10–100x faster than pip
> - Creates and manages the virtual environment for you
> - `uv.lock` guarantees identical installs across dev, staging, and production
> - No need to manually activate the venv for most commands — prefix with `uv run`

---

### 4. Activate the virtual environment

```bash
# macOS / Linux
source .venv/bin/activate

# Windows
.venv\Scripts\activate
```

Or skip activation entirely and prefix commands with `uv run` (see [Running the App](#running-the-app)).

---

### 5. Set up environment variables

```bash
cp .env.example .env
```

Open `.env` and fill in your Anthropic API key:

```env
ANTHROPIC_API_KEY=sk-ant-...
```

See [Configuration](#configuration) for all available variables.

---

## Configuration

All settings are loaded from `.env` via `backend/app/config.py`.

| Variable | Default | Required | Description |
|---|---|---|---|
| `ANTHROPIC_API_KEY` | — | **Yes** | Your Anthropic API key |
| `ANTHROPIC_MODEL` | `claude-sonnet-4-6` | No | Primary model for extraction |
| `ANTHROPIC_FALLBACK_MODEL` | `claude-opus-4-6` | No | Used if the primary model fails |
| `BACKEND_BASE_URL` | `http://localhost:8000` | No | Used to build download URLs in API responses |
| `STORAGE_DIR` | `storage` | No | Directory where job files are written |
| `MAX_FILE_SIZE_MB` | `32` | No | Maximum PDF upload size in MB |
| `MAX_PDF_PAGES` | `100` | No | Maximum number of pages per PDF |
| `TARGET_LANGUAGE` | `English` | No | Language to translate into |

---

## Running the App

You need **two terminals** — one for the backend, one for the frontend.

### With virtual environment activated

**Terminal 1 — Backend:**
```bash
uvicorn backend.app.main:app --reload --port 8000
```

**Terminal 2 — Frontend:**
```bash
streamlit run frontend/streamlit_app.py
```

---

### With uv run (no activation needed)

**Terminal 1 — Backend:**
```bash
uv run uvicorn backend.app.main:app --reload --port 8000
```

**Terminal 2 — Frontend:**
```bash
uv run streamlit run frontend/streamlit_app.py
```

---

Open the app at **http://localhost:8501**

The backend API is available at **http://localhost:8000**

Interactive API docs (auto-generated by FastAPI): **http://localhost:8000/docs**

---

## API Reference

### `POST /api/v1/translate`

Upload a PDF and receive the full translation.

**Request:** `multipart/form-data`
- `file` — PDF file (required)

**Response:** JSON
```json
{
  "job_id": "3f9a1c2e-...",
  "filename": "claim.pdf",
  "pages": 5,
  "model_used": "claude-sonnet-4-6",
  "fallback_used": false,
  "extracted_document": {
    "title": "Translated Medical Claim Document",
    "source_language_guess": "Arabic / English mixed",
    "target_language": "English",
    "sections": [ ... ],
    "warnings": [ ... ]
  },
  "pdf_download_url": "http://localhost:8000/api/v1/download/3f9a1c2e-.../translated_document.pdf",
  "json_download_url": "http://localhost:8000/api/v1/download/3f9a1c2e-.../translated_document.json"
}
```

**Error codes:**
- `400` — file is not a PDF or is empty
- `500` — pipeline failure (size limit, page limit, Claude error, render error)

---

### `GET /api/v1/download/{job_id}/{filename}`

Download a file from a completed job.

**Response:** file stream
- `application/pdf` for `.pdf` files
- `application/json` for `.json` files

**Error codes:**
- `404` — job ID or filename not found

---

### `GET /api/v1/health`

```json
{ "status": "ok" }
```

---

## Design Decisions

### One LLM step — extraction and translation together
Claude handles both extraction and translation in a single API call. The PDF is sent as a base64-encoded `document` block which Claude reads natively — no OCR library, no pre-processing. This means fewer moving parts, faster results, and one single point of failure to debug.

### Code-based PDF rendering — not LLM-generated layout
The translated PDF is built deterministically by `pdf_renderer.py` using ReportLab. Asking a second LLM to design the layout was deliberately avoided because:
- Stable, consistent formatting on every run
- Easy to change branding or styles in one place
- No hallucinated layout decisions
- Predictable cost

### Primary + fallback model chain
The pipeline tries `claude-sonnet-4-6` first (fast, cost-effective). If it fails for any reason — network error, malformed JSON, Pydantic validation failure — it automatically retries with `claude-opus-4-6` (more capable, better with difficult handwriting and low-quality scans). If both fail, a `500` is returned with a combined error message.

### Strict Pydantic validation as a contract
Claude's output is validated against a strict schema before anything is written to disk or rendered. Fields with unexpected types or values trigger the fallback instead of silently producing a broken output. Coercion validators handle minor model inconsistencies (e.g. `"error"` → `"critical"`, `"warn"` → `"warning"`) without breaking the pipeline.

### Job-based UUID storage
Each translation gets a UUID folder under `storage/`. This ensures:
- No filename collisions between concurrent jobs
- Original and output files are grouped together
- Download URLs are stable and unique
- Easy to inspect or debug any past job by browsing the folder

---

## Production Checklist

The current codebase is a solid MVP. Before going to production, consider adding:

- [ ] Authentication — protect the `/translate` endpoint
- [ ] Background job queue — avoid HTTP timeout on large PDFs (Celery, ARQ, or FastAPI background tasks)
- [ ] Object storage — replace local `storage/` with S3 or equivalent
- [ ] Database — track job history, status, and user association
- [ ] Retry logic — automatic retry with backoff on Claude API errors
- [ ] Confidence-based review queue — flag low-confidence documents for human review
- [ ] Prompt caching — cache system prompt for repeated document types
- [ ] Page chunking — split very large PDFs and merge results
- [ ] Rate limiting — protect the API from abuse
- [ ] Structured logging — replace `print()` with `logging` or a structured logger
