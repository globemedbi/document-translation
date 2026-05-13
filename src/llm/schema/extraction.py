from __future__ import annotations

from typing import Optional

from pydantic import BaseModel, Field


class BBoxSchema(BaseModel):
    x: float = 0.0
    y: float = 0.0
    w: float = 0.1
    h: float = 0.02


class FormFieldSchema(BaseModel):
    key: str
    value: str = ""
    field_type: str = "text"
    key_bbox: Optional[BBoxSchema] = None
    value_bbox: Optional[BBoxSchema] = None
    confidence: str = "high"
    notes: str = ""


class PageExtractionResponse(BaseModel):
    page_number: int
    source_language: str = ""
    document_type: str = "form"
    fields: list[FormFieldSchema] = Field(default_factory=list)


class DocInfoResponse(BaseModel):
    source_language: str = "Unknown"
    title: str = "Document"
    document_type: str = "form"
