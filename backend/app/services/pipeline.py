from __future__ import annotations

import uuid
from pathlib import Path

import fitz

from backend.app.config import settings
from backend.app.models.schemas import DocumentWarning, ExtractedDocument, TranslateResponse
from backend.app.services.anthropic_client import make_fallback_client, make_primary_client
from backend.app.services.extraction_prompt import (
    build_extraction_system_prompt,
    build_user_prompt,
)
from backend.app.services.pdf_renderer import PdfRenderer
from backend.app.utils.file_store import JobStore


class DocumentPipeline:
    def __init__(self) -> None:
        self.store = JobStore(settings.storage_path)
        self.renderer = PdfRenderer()
        self.primary_client = make_primary_client()
        self.fallback_client = make_fallback_client()

    def process(self, *, filename: str, pdf_bytes: bytes) -> TranslateResponse:
        print(f"[pipeline] Received file: {filename} ({len(pdf_bytes) / 1024:.1f} KB)")
        try:
            self._validate_pdf_size(pdf_bytes)
        except ValueError as exc:
            print(f"[pipeline] ERROR: File size validation failed — {exc}")
            raise
        try:
            pages = self._count_pages(pdf_bytes)
        except Exception as exc:
            print(f"[pipeline] ERROR: Could not open or read PDF — {exc}")
            raise ValueError(f"Invalid or corrupted PDF: {exc}") from exc
        print(f"[pipeline] Page count: {pages}")
        if pages > settings.max_pdf_pages:
            raise ValueError(
                f"PDF has {pages} pages, which exceeds the configured limit of {settings.max_pdf_pages}."
            )

        job_id = str(uuid.uuid4())
        print(f"[pipeline] Job ID: {job_id}")
        self.store.save_upload(job_id, filename, pdf_bytes)

        system_prompt = build_extraction_system_prompt(settings.target_language)
        user_prompt = build_user_prompt()

        fallback_used = False
        model_used: str = ""
        document: ExtractedDocument | None = None
        errors: list[str] = []

        for client, is_fallback in [
            (self.primary_client, False),
            (self.fallback_client, True),
        ]:
            print(f"[pipeline] Calling model: {client.model} (fallback={is_fallback})")
            try:
                raw_result = client.extract_document(
                    pdf_bytes=pdf_bytes,
                    system_prompt=system_prompt,
                    user_prompt=user_prompt,
                )
                document = ExtractedDocument.model_validate(raw_result)
                fallback_used = is_fallback
                model_used = client.model
                print(f"[pipeline] Extraction succeeded with model: {model_used}")
                break
            except Exception as exc:  # broad on purpose for MVP fallback
                print(f"[pipeline] Model {client.model} failed: {exc}")
                errors.append(f"{client.model}: {exc}")
        else:
            raise RuntimeError(
                "Document extraction failed for both primary and fallback models. "
                + " | ".join(errors)
            )

        if errors and document is not None:
            for err in errors:
                document.warnings.append(
                    DocumentWarning(page_number=None, message=f"Fallback note: {err}", severity="info")
                )
            document = ExtractedDocument.model_validate(document.model_dump())

        json_filename = "translated_document.json"
        pdf_filename = "translated_document.pdf"

        assert document is not None  # guaranteed: for/else raises before reaching here

        print(f"[pipeline] Saving outputs for job {job_id}")
        try:
            self.store.save_json(job_id, json_filename, document.model_dump())
        except Exception as exc:
            print(f"[pipeline] ERROR: Failed to save JSON for job {job_id} — {exc}")
            raise
        try:
            output_pdf_path = self.store.job_dir(job_id) / pdf_filename
            self.renderer.render(document, output_pdf_path)
        except Exception as exc:
            print(f"[pipeline] ERROR: Failed to render PDF for job {job_id} — {exc}")
            raise
        print(f"[pipeline] Done. PDF: {output_pdf_path}")

        return TranslateResponse(
            job_id=job_id,
            filename=filename,
            pages=pages,
            model_used=model_used,
            fallback_used=fallback_used,
            extracted_document=document,
            pdf_download_url=f"{settings.backend_base_url}/api/v1/download/{job_id}/{pdf_filename}",
            json_download_url=f"{settings.backend_base_url}/api/v1/download/{job_id}/{json_filename}",
        )

    @staticmethod
    def _count_pages(pdf_bytes: bytes) -> int:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
            return document.page_count

    @staticmethod
    def _validate_pdf_size(pdf_bytes: bytes) -> None:
        size_mb = len(pdf_bytes) / (1024 * 1024)
        if size_mb > settings.max_file_size_mb:
            raise ValueError(
                f"PDF size is {size_mb:.2f} MB, above the configured limit of {settings.max_file_size_mb} MB."
            )
