from __future__ import annotations

from pydantic import BaseModel, Field


class TranslationItem(BaseModel):
    translated_key: str
    translated_value: str


class TranslationBatchResponse(BaseModel):
    items: list[TranslationItem] = Field(default_factory=list)
