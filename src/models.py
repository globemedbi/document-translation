from __future__ import annotations

from typing import Optional
from pydantic import BaseModel, Field


class BBox(BaseModel):
    """Bounding box as fractions of page size. x/y = top-left corner."""
    x: float = Field(ge=0.0, le=1.0, description="Left edge, 0..1")
    y: float = Field(ge=0.0, le=1.0, description="Top edge, 0..1")
    w: float = Field(ge=0.0, le=1.0, description="Width, 0..1")
    h: float = Field(ge=0.0, le=1.0, description="Height, 0..1")


class FormField(BaseModel):
    key: str                        # Field label / name
    value: str                      # Filled-in content (empty string if blank)
    field_type: str = "text"        # text | checkbox | date | header | signature | table
    key_bbox: Optional[BBox] = None
    value_bbox: Optional[BBox] = None
    page: int
    confidence: str = "high"        # high | medium | low
    notes: str = ""


class PageExtraction(BaseModel):
    page_number: int
    fields: list[FormField]
    width_px: int = 0
    height_px: int = 0


class DocumentExtraction(BaseModel):
    pages: list[PageExtraction]
    source_language: str = "Unknown"
    title: str = "Document"
    document_type: str = "form"
    total_pages: int = 0


class TranslatedField(BaseModel):
    original: FormField
    translated_key: str
    translated_value: str


class TranslatedPage(BaseModel):
    page_number: int
    width_px: int = 0
    height_px: int = 0
    translated_fields: list[TranslatedField]


class TranslatedDocument(BaseModel):
    original: DocumentExtraction
    pages: list[TranslatedPage]
    target_language: str
