# PDF Form Translator

Translates medical and insurance PDF forms from any language into any target language, then reconstructs the document with the translated content — preserving layout, structure, and visual fidelity.

Built on Claude (Anthropic) for extraction, translation, and document reconstruction. Supports four reconstruction modes with different speed/quality trade-offs.

---

## Quickstart (Docker)

```bash
git clone <repo>
cd v3
cp .env.example .env
# Add your Anthropic API key to .env
docker compose up --build
```

| Service | URL |
|---|---|
| Streamlit UI | http://localhost:8501 |
| FastAPI backend | http://localhost:8000 |
| API docs | http://localhost:8000/docs |

---

## Modes

| Mode | Speed | Quality | Best For |
|---|---|---|---|
| **AUTO** | — | — | Let the system decide (TEXT for digital PDFs, VISION for scanned) |
| **TEXT** | ⚡ ~30s | ★★★ | PDFs with selectable text (Word exports, digital forms) |
| **VISION** | 🔵 ~2-3 min | ★★★★ | Any PDF — overlays translated text on the original |
| **CODING** | 🐢 ~5-10 min | ★★★★★ | Complex forms — fully recreates the document as an editable DOCX |

### TEXT mode
Extracts embedded text spans directly with PyMuPDF (no Vision API call), translates them in parallel batches, then redacts the originals and re-inserts the translations in-place. Fastest, but only works when the PDF has selectable text.

### VISION mode
Claude Vision reads each page image and extracts every visible field with its bounding box. Translated text is drawn directly on the PDF using PyMuPDF — white rectangle over the original, translated text inserted at the same position with matching font size and RTL support.

### CODING mode
A per-page Plan → Action → Observe loop (all pages run in parallel):
1. **PLAN** — Claude analyses the page image and produces a precise layout plan
2. **ACTION** — Claude writes a complete `python-docx` script for that page
3. **Execute** — script is syntax-checked then run in a subprocess
4. **OBSERVE** — Claude scores the output against the original (1–10). If < 8, the feedback is fed back to ACTION for a retry

All per-page DOCX files are merged into a single document. On macOS with Microsoft Word installed, the result is also converted to PDF.

---

## Architecture

```
PDF Input
   │
   ├─── TEXT mode ──────────────────────────────────────────────────────────────
   │    PyMuPDF extracts text spans → parallel batch translation → redact + reinsert
   │
   └─── VISION / CODING mode ───────────────────────────────────────────────────
        │
        ▼
   ┌─────────────────────────────────────────┐
   │  Step 1 — EXTRACT  (src/extractor.py)  │
   │  Claude Vision: page image → JSON      │
   │  fields with bounding boxes            │
   └─────────────────────────────────────────┘
        │
        ▼
   ┌─────────────────────────────────────────┐
   │  Step 2 — TRANSLATE  (src/translator)  │
   │  Claude: field labels + values →       │
   │  translated equivalents                │
   └─────────────────────────────────────────┘
        │
        ├── VISION ──────────────────────────────────────────────────────────────
        │   PyMuPDF: white rect over original + translated text overlay
        │
        └── CODING ──────────────────────────────────────────────────────────────
            PAO loop per page (parallel, ≤3 at once)
            PLAN → ACTION → OBSERVE → retry until score ≥ 8
            Merge pages → DOCX (→ PDF if Word available)
```

---

## Local Setup

**Prerequisites:** Python 3.11+, [uv](https://docs.astral.sh/uv/getting-started/installation/), Anthropic API key.

```bash
uv sync
cp .env.example .env
# Edit .env: set ANTHROPIC_API_KEY
```

### CLI

```bash
# Auto-detect mode
uv run python main.py form.pdf

# Explicit mode and language
uv run python main.py form.pdf --mode VISION --language French
uv run python main.py form.pdf --mode TEXT --language Arabic
uv run python main.py form.pdf --mode CODING --language Spanish --output out.docx
```

### API only

```bash
uv run uvicorn api:app --port 8000
```

### Streamlit + API

```bash
uv run uvicorn api:app --port 8000 &
uv run streamlit run app.py
```

---

## API Reference

### `POST /translate`

Upload a PDF and receive the translated file.

| Field | Type | Default | Description |
|---|---|---|---|
| `file` | file | — | PDF to translate |
| `language` | string | `French` | Target language name (e.g. `Arabic`, `Spanish`) |
| `mode` | string | `AUTO` | `AUTO` \| `TEXT` \| `VISION` \| `CODING` |

**Response:** The translated file (`application/pdf` or `.docx`) as a direct download.

**Example (curl):**
```bash
curl -X POST http://localhost:8000/translate \
  -F "file=@form.pdf" \
  -F "language=French" \
  -F "mode=AUTO" \
  --output translated.pdf
```

**Example (Python):**
```python
import httpx

with open("form.pdf", "rb") as f:
    resp = httpx.post(
        "http://localhost:8000/translate",
        data={"language": "French", "mode": "AUTO"},
        files={"file": ("form.pdf", f, "application/pdf")},
        timeout=600.0,
    )

resp.raise_for_status()
with open("translated.pdf", "wb") as out:
    out.write(resp.content)
```

### `GET /health`

Returns `{"status": "ok"}` — used by Docker health checks.

---

## Configuration

All settings are read from environment variables (or `.env`):

```env
# Required
ANTHROPIC_API_KEY=sk-ant-...

# Models
ANTHROPIC_MODEL=claude-sonnet-4-6         # fast model: translation, TEXT mode, coding
ANTHROPIC_HEAVY_MODEL=claude-opus-4-7     # heavy model: extraction (Vision)

# Translation
TARGET_LANGUAGE=French                    # default language (overridable per request)
SOLUTION_MODE=AUTO                        # AUTO | TEXT | VISION | CODING

# CODING mode
MAX_AGENT_ITERATIONS=3                    # PAO retries per page (higher = better, slower)
AGENT_WORKSPACE=workspace                 # temp directory for generated scripts

# Output
OUTPUT_DIR=output
LOG_LEVEL=INFO
```

---

## Project Structure

```
├── main.py                         # CLI entry point
├── api.py                          # FastAPI backend (POST /translate)
├── app.py                          # Streamlit web frontend
├── Dockerfile
├── docker-compose.yml
├── pyproject.toml
├── uv.lock
└── src/
    ├── config.py                   # Settings from environment
    ├── models.py                   # Pydantic data models
    ├── extractor.py                # Step 1: PDF image → structured fields
    ├── translator.py               # Step 2: fields → translated fields
    ├── pdf_utils.py                # PDF rendering and image helpers
    └── solutions/
        ├── text_layer.py           # TEXT mode: direct span replacement
        ├── vision_overlay.py       # VISION mode: PyMuPDF overlay
        └── coding_agent.py         # CODING mode: PAO agent loop
```

---

## Output

| Mode | Output location | Format |
|---|---|---|
| TEXT | `output/text/<stem>_translated.pdf` | PDF |
| VISION | `output/vision/<stem>_translated.pdf` | PDF |
| CODING | `output/coding/<stem>_translated.docx` | DOCX (+ PDF if Word is installed) |

Logs are written to `<output_dir>/translation.log` (DEBUG level) and stderr (INFO level).

---

## Notes

- **DOCX → PDF** in CODING mode requires Microsoft Word on macOS/Windows. Without it, the `.docx` is returned and a skip message is logged — no error.
- **Scanned PDFs** (no embedded text): TEXT mode is not applicable; VISION/CODING modes use Claude's estimated bounding boxes which may be slightly less precise than native coordinates.
- **Complex pages** (50+ fields): CODING mode uses the best result seen across all PAO iterations even if the target score is not reached.
- The `ANTHROPIC_API_KEY` is read server-side from `.env` — it is never required or accepted in API requests.
