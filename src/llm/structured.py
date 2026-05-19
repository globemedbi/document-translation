from __future__ import annotations

import base64
import io
from typing import Any, TypeVar

from PIL import Image
from pydantic import BaseModel

T = TypeVar("T", bound=BaseModel)


def _is_openai(client: Any) -> bool:
    return type(client).__module__.startswith("openai")


def image_block(client: Any, img: Image.Image) -> dict:
    """Build a provider-correct image content block from a PIL Image."""
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    if _is_openai(client):
        return {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
    return {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}}


def image_block_from_bytes(client: Any, data: bytes) -> dict:
    """Build a provider-correct image content block from raw bytes."""
    b64 = base64.b64encode(data).decode()
    if _is_openai(client):
        return {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
    return {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}}


def chat(
    client: Any,
    model: str,
    messages: list[dict[str, Any]],
    *,
    system: str | None = None,
    max_tokens: int = 4096,
    reasoning_effort: str | None = None,
) -> str:
    """Send a chat request and return the text response."""
    if _is_openai(client):
        full = []
        if system:
            full.append({"role": "system", "content": system})
        full.extend(messages)
        kwargs: dict[str, Any] = {"model": model, "max_completion_tokens": max_tokens, "messages": full}
        if reasoning_effort:
            kwargs["reasoning_effort"] = reasoning_effort
        resp = client.chat.completions.create(**kwargs)
        return (resp.choices[0].message.content or "").strip()
    else:
        kwargs = {"model": model, "max_tokens": max_tokens, "messages": messages}
        if system:
            kwargs["system"] = system
        resp = client.messages.create(**kwargs)
        return resp.content[0].text.strip()


def structured_call(
    client: Any,
    model: str,
    messages: list[dict[str, Any]],
    schema: type[T],
    *,
    system: str | None = None,
    max_tokens: int = 4096,
    reasoning_effort: str | None = None,
) -> T:
    """Call the LLM with structured output and return a validated Pydantic instance."""
    if _is_openai(client):
        full: list[dict[str, Any]] = []
        if system:
            full.append({"role": "system", "content": system})
        full.extend(messages)
        kwargs: dict[str, Any] = {
            "model": model, "max_completion_tokens": max_tokens, "messages": full, "response_format": schema,
        }
        if reasoning_effort:
            kwargs["reasoning_effort"] = reasoning_effort
        resp = client.beta.chat.completions.parse(**kwargs)
        result = resp.choices[0].message.parsed
        if result is None:
            raise ValueError(f"No structured output for {schema.__name__} (finish_reason={resp.choices[0].finish_reason})")
        return result
    else:
        kwargs = {
            "model": model, "max_tokens": max_tokens, "messages": messages, "output_format": schema,
        }
        if system:
            kwargs["system"] = system
        resp = client.messages.parse(**kwargs)
        result = resp.parsed_output
        if result is None:
            raise ValueError(f"No structured output for {schema.__name__} (stop_reason={resp.stop_reason})")
        return result
