from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from loguru import logger
from PIL import Image

from src.config import get_client, settings
from src.llm.schema.extraction import BBoxSchema, DocInfoResponse, FormFieldSchema, PageExtractionResponse
from src.llm.structured import image_block, structured_call
from src.models import BBox, DocumentExtraction, FormField, PageExtraction
from src.pdf_utils import pdf_to_images

# ── Prompts ───────────────────────────────────────────────────────────────────

_EXTRACTION_SYSTEM = """\
You are an expert document analyst specialised in bilingual (Arabic/English) medical and
insurance forms. Your job is to extract EVERY piece of visible text from the provided page
image as structured key-value pairs.

Rules:
- Capture all printed labels AND all handwritten or typed fill-in values.
- For each field provide the label as "key" and the filled content as "value".
- For redacted / blacked-out areas use value "[REDACTED]".
- For blank fields use value "".
- For section titles / headers use field_type "header" and put the title in both key and value.
- For checkboxes include whether they are checked: value "☑ checked" or "☐ unchecked".
- Provide approximate bounding boxes as fractions of the page (0.0 = left/top, 1.0 = right/bottom).
"""

_EXTRACTION_USER = """\
Extract all form fields from this page image.

For each field provide:
- key: the field label text
- value: the filled-in content (empty string if blank, "[REDACTED]" if blacked out)
- field_type: text | checkbox | date | header | signature | table
- key_bbox / value_bbox: bounding box as fractions of page size (x, y, w, h in 0..1)
- confidence: high | medium | low
- notes: any relevant observation (optional)

Set page_number to the page being extracted.
Set source_language to the detected language(s).
Set document_type to: receipt | claim_form | approval_form | id_document | other
If a field has no separate value area (e.g. a pure header) set value_bbox to null.
"""

_DOC_INFO_USER = """\
Look at this page image and identify the document metadata:
- source_language: the language(s) used in the document
- title: the document's title or name
- document_type: receipt | claim_form | approval_form | id_document | other
"""


def _safe_bbox(raw: BBoxSchema | None) -> BBox | None:
    if not raw:
        return None
    try:
        return BBox(
            x=max(0.0, min(1.0, raw.x)),
            y=max(0.0, min(1.0, raw.y)),
            w=max(0.001, min(1.0, raw.w)),
            h=max(0.001, min(1.0, raw.h)),
        )
    except Exception:
        return None


# ── Extractor class ───────────────────────────────────────────────────────────

class Extractor:
    def __init__(self) -> None:
        self.client = get_client()
        self.model = settings.model
        logger.info(f"Extractor initialised (provider={settings.llm_provider}, model={self.model})")

    # ── Public ────────────────────────────────────────────────────────────────

    def extract_document(self, pdf_path: Path) -> DocumentExtraction:
        logger.info(f"=== EXTRACTION START: {pdf_path.name} ===")
        page_images = pdf_to_images(pdf_path)

        doc_info = self._get_doc_info(page_images[0][1]) if page_images else DocInfoResponse()

        n = len(page_images)
        page_extractions: list[PageExtraction | None] = [None] * n

        with ThreadPoolExecutor(max_workers=min(n, 3)) as pool:
            futures = {
                pool.submit(self._extract_page, page_num, img): (idx, page_num)
                for idx, (page_num, img) in enumerate(page_images)
            }
            for future in as_completed(futures):
                idx, page_num = futures[future]
                try:
                    pe = future.result()
                    page_extractions[idx] = pe
                    logger.info(f"  page {page_num}: {len(pe.fields)} fields extracted")
                except Exception as exc:
                    logger.error(f"  page {page_num}: extraction failed — {exc}")
                    w, h = page_images[idx][1].size
                    page_extractions[idx] = PageExtraction(
                        page_number=page_num, fields=[], width_px=w, height_px=h
                    )

        extractions: list[PageExtraction] = [p for p in page_extractions if p is not None]
        total_fields = sum(len(p.fields) for p in extractions)
        doc = DocumentExtraction(
            pages=extractions,
            source_language=doc_info.source_language,
            title=doc_info.title,
            document_type=doc_info.document_type,
            total_pages=n,
        )
        logger.info(
            f"=== EXTRACTION DONE: {len(extractions)} pages, "
            f"{total_fields} total fields ==="
        )
        return doc

    # ── Private ───────────────────────────────────────────────────────────────

    def _get_doc_info(self, img: Image.Image) -> DocInfoResponse:
        logger.debug("Getting document-level metadata from page 1 …")
        try:
            return structured_call(
                self.client,
                self.model,
                messages=[{
                    "role": "user",
                    "content": [
                        self._image_block(img),
                        {"type": "text", "text": _DOC_INFO_USER},
                    ],
                }],
                schema=DocInfoResponse,
                max_tokens=512,
                reasoning_effort="medium",
            )
        except Exception as exc:
            logger.warning(f"Could not get doc info: {exc}")
            return DocInfoResponse()

    def _extract_page(self, page_num: int, img: Image.Image) -> PageExtraction:
        for attempt in range(1, 3):
            try:
                result = structured_call(
                    self.client,
                    self.model,
                    messages=[{
                        "role": "user",
                        "content": [
                            self._image_block(img),
                            {"type": "text", "text": _EXTRACTION_USER},
                        ],
                    }],
                    schema=PageExtractionResponse,
                    system=_EXTRACTION_SYSTEM,
                    max_tokens=16000,
                    reasoning_effort="medium",
                )
                logger.debug(f"Page {page_num} attempt {attempt}: {len(result.fields)} fields")
                fields = self._parse_fields(result.fields, page_num)
                return PageExtraction(
                    page_number=page_num,
                    fields=fields,
                    width_px=img.width,
                    height_px=img.height,
                )
            except Exception as exc:
                logger.warning(
                    f"Page {page_num} attempt {attempt}: structured output error ({exc}). "
                    f"{'Retrying …' if attempt == 1 else 'Giving up.'}"
                )
                if attempt == 1:
                    continue
                logger.error(f"Page {page_num}: returning empty extraction after failure")
                return PageExtraction(
                    page_number=page_num,
                    fields=[],
                    width_px=img.width,
                    height_px=img.height,
                )

        return PageExtraction(page_number=page_num, fields=[], width_px=img.width, height_px=img.height)

    def _parse_fields(self, raw_fields: list[FormFieldSchema], page_num: int) -> list[FormField]:
        fields: list[FormField] = []
        for i, f in enumerate(raw_fields):
            try:
                fields.append(FormField(
                    key=f.key.strip(),
                    value=f.value.strip(),
                    field_type=f.field_type,
                    key_bbox=_safe_bbox(f.key_bbox),
                    value_bbox=_safe_bbox(f.value_bbox),
                    page=page_num,
                    confidence=f.confidence,
                    notes=f.notes,
                ))
            except Exception as exc:
                logger.warning(f"Skipping malformed field {i} on page {page_num}: {exc}")
        return fields

    def _image_block(self, img: Image.Image) -> dict:
        return image_block(self.client, img)
