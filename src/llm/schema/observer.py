from __future__ import annotations

from pydantic import BaseModel, Field


class ObserverResult(BaseModel):
    score: int = Field(default=5, ge=1, le=10)
    is_done: bool = False
    missing_fields: list[str] = Field(default_factory=list)
    wrong_values: list[str] = Field(default_factory=list)
    feedback: str = ""
