from __future__ import annotations

from pydantic import BaseModel, Field


class TextTranslationItem(BaseModel):
    original: str
    translated: str


class TextTranslationResponse(BaseModel):
    translations: list[TextTranslationItem] = Field(default_factory=list)
