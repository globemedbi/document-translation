from __future__ import annotations

from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from backend.app.config import settings
from backend.app.models.schemas import TranslateResponse
from backend.app.services.pipeline import DocumentPipeline

router = APIRouter(prefix="/api/v1", tags=["translator"])
pipeline = DocumentPipeline()


@router.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@router.post("/translate", response_model=TranslateResponse)
async def translate_pdf(file: UploadFile = File(...)) -> TranslateResponse:
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    pdf_bytes = await file.read()
    if not pdf_bytes:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    try:
        return pipeline.process(filename=file.filename, pdf_bytes=pdf_bytes)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@router.get("/download/{job_id}/{filename}")
def download_file(job_id: str, filename: str):
    path = settings.storage_path / job_id / filename
    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")
    media_type = "application/pdf" if path.suffix.lower() == ".pdf" else "application/json"
    return FileResponse(path=path, filename=filename, media_type=media_type)
