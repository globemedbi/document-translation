"""
PDF Translator — FastAPI backend.

Single endpoint: POST /translate
  - Accepts a PDF + language + mode as multipart form data
  - API key is read from environment (OPENAI_API_KEY or ANTHROPIC_API_KEY) — not required in the request
  - Runs the pipeline synchronously and returns the translated file directly

Run with:
    uv run uvicorn api:app --port 8000
"""

from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

load_dotenv()

app = FastAPI(title="PDF Translator API", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

_ROOT = Path(__file__).parent


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/translate")
async def translate(
    file: UploadFile = File(...),
    language: str = Form("French"),
    mode: str = Form("AUTO"),
) -> FileResponse:
    """
    Upload a PDF and get back the translated file.
    API key is read from OPENAI_API_KEY or ANTHROPIC_API_KEY environment variable.
    mode: AUTO | TEXT | VISION | CODING
    """
    provider = os.environ.get("LLM_PROVIDER", "anthropic").lower()
    if provider == "openai" and not os.environ.get("OPENAI_API_KEY"):
        raise HTTPException(500, "OPENAI_API_KEY not set.")
    if provider != "openai" and not os.environ.get("ANTHROPIC_API_KEY"):
        raise HTTPException(500, "ANTHROPIC_API_KEY not set.")

    fname = file.filename or "document.pdf"
    if not fname.lower().endswith(".pdf"):
        raise HTTPException(400, "Only PDF files are accepted.")

    pdf_bytes = await file.read()
    if len(pdf_bytes) < 100:
        raise HTTPException(400, "Uploaded file appears to be empty.")

    mode = mode.upper()
    stem = re.sub(r"[^\w\-]", "_", Path(fname).stem)

    print(f"\n{'='*60}")
    print(f"  NEW JOB: {fname}")
    print(f"  Language : {language}")
    print(f"  Mode     : {mode}")
    print(f"{'='*60}")

    # Resolve AUTO before creating the folder so output lands in the right place
    if mode == "AUTO":
        from src.pdf_utils import is_text_pdf as _is_text
        tmp_path = _ROOT / "output" / f"_tmp_{stem}.pdf"
        tmp_path.write_bytes(pdf_bytes)
        resolved_mode = "TEXT" if _is_text(tmp_path) else "VISION"
        tmp_path.unlink(missing_ok=True)
        print(f"  AUTO → detected mode: {resolved_mode}")
    else:
        resolved_mode = mode

    job_dir = _ROOT / "output" / resolved_mode.lower() / stem
    job_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = job_dir / f"{stem}.pdf"
    pdf_path.write_bytes(pdf_bytes)

    out_ext = ".docx" if resolved_mode == "CODING" else ".pdf"
    output_path = job_dir / f"{stem}_translated{out_ext}"

    print(f"  Output   : {output_path}")
    print(f"  Running pipeline …\n")

    result = subprocess.run(
        [
            "uv", "run", "python", "main.py",
            str(pdf_path),
            "--mode", resolved_mode,
            "--language", language,
            "--output", str(output_path),
        ],
        cwd=str(_ROOT),
        env={**os.environ, "OUTPUT_DIR": str(job_dir), "LOG_LEVEL": "INFO"},
    )

    if result.returncode != 0:
        print(f"  [ERROR] Pipeline failed (exit code {result.returncode})")
        log_file = job_dir / "translation.log"
        detail = log_file.read_text(encoding="utf-8")[-2000:] if log_file.exists() else "Pipeline failed"
        raise HTTPException(500, detail)

    actual = _find_output(job_dir, stem)
    if not actual:
        print(f"  [ERROR] Output file not found in {job_dir}")
        raise HTTPException(500, "Pipeline completed but output file was not found.")

    print(f"\n  ✓ Done → {actual}")
    print(f"{'='*60}\n")

    mime = (
        "application/pdf" if actual.suffix.lower() == ".pdf"
        else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return FileResponse(str(actual), media_type=mime, filename=actual.name)


def _find_output(job_dir: Path, stem: str) -> Path | None:
    for ext in (".pdf", ".docx"):
        p = job_dir / f"{stem}_translated{ext}"
        if p.exists():
            return p
    for pattern in ("*.pdf", "*.docx"):
        hits = [f for f in job_dir.rglob(pattern) if "_translated" in f.name]
        if hits:
            return hits[0]
    return None
