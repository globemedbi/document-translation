from __future__ import annotations

from typing import Any, TypeVar

import anthropic
from pydantic import BaseModel

T = TypeVar("T", bound=BaseModel)


def structured_call(
    client: anthropic.Anthropic,
    model: str,
    messages: list[dict[str, Any]],
    schema: type[T],
    *,
    system: str | None = None,
    max_tokens: int = 4096,
) -> T:
    """
    Call the Anthropic API with structured output and return a validated Pydantic instance.
    Uses client.messages.parse() with output_format.
    """
    kwargs: dict[str, Any] = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": messages,
        "output_format": schema,
    }
    if system:
        kwargs["system"] = system

    resp = client.messages.parse(**kwargs)
    result = resp.parsed_output
    if result is None:
        raise ValueError(
            f"No structured output returned for {schema.__name__} "
            f"(stop_reason={resp.stop_reason})"
        )
    return result
