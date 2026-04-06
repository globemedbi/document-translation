from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field, field_validator


ConfidenceLevel = Literal["high", "medium", "low"]

_SEVERITY_COERCE: dict[str, str] = {
    "error": "critical",
    "err": "critical",
    "warn": "warning",
    "note": "info",
    "debug": "info",
}

_CONFIDENCE_COERCE: dict[str, str] = {
    "unknown": "low",
    "uncertain": "low",
    "sure": "high",
    "confident": "high",
}


class DocumentField(BaseModel):
    label_en: str = Field(..., description="Field label translated to English")
    value_en: str = Field(..., description="Field value translated or normalized to English")
    value_original: str | None = Field(
        default=None,
        description="Original visible value before translation, if available",
    )
    page_number: int = Field(..., ge=1)
    confidence: ConfidenceLevel = "medium"
    notes: str | None = None

    @field_validator("confidence", mode="before")
    @classmethod
    def coerce_confidence(cls, v: object) -> object:
        if isinstance(v, str):
            return _CONFIDENCE_COERCE.get(v.lower().strip(), v)
        return v


class DocumentTableRow(BaseModel):
    cells: list[str]
    page_number: int = Field(..., ge=1)
    confidence: ConfidenceLevel = "medium"

    @field_validator("confidence", mode="before")
    @classmethod
    def coerce_confidence(cls, v: object) -> object:
        if isinstance(v, str):
            return _CONFIDENCE_COERCE.get(v.lower().strip(), v)
        return v


class DocumentTable(BaseModel):
    title: str
    columns: list[str]
    rows: list[DocumentTableRow] = Field(default_factory=list)


class DocumentSection(BaseModel):
    title: str
    page_numbers: list[int] = Field(default_factory=list)
    summary: str | None = None
    fields: list[DocumentField] = Field(default_factory=list)
    tables: list[DocumentTable] = Field(default_factory=list)


class DocumentWarning(BaseModel):
    page_number: int | None = None
    message: str
    severity: Literal["info", "warning", "critical"] = "warning"

    @field_validator("severity", mode="before")
    @classmethod
    def coerce_severity(cls, v: object) -> object:
        if isinstance(v, str):
            return _SEVERITY_COERCE.get(v.lower().strip(), v)
        return v


class ExtractedDocument(BaseModel):
    source_language_guess: str | None = None
    target_language: str = "English"
    document_type: str = "medical_claim_package"
    title: str = "Translated Medical Claim Document"
    sections: list[DocumentSection] = Field(default_factory=list)
    warnings: list[DocumentWarning] = Field(default_factory=list)


class TranslateResponse(BaseModel):
    job_id: str
    filename: str
    pages: int
    model_used: str
    fallback_used: bool = False
    extracted_document: ExtractedDocument
    pdf_download_url: str
    json_download_url: str
