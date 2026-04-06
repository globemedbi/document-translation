from __future__ import annotations

import base64
from typing import Any

import httpx

from backend.app.config import settings
from backend.app.utils.json_utils import extract_json_object


class AnthropicClient:
    BASE_URL = "https://api.anthropic.com/v1/messages"

    def __init__(self, api_key: str | None, model: str, timeout_seconds: int = 180):
        self.api_key = api_key
        self.model = model
        self.timeout_seconds = timeout_seconds

    def _build_headers(self) -> dict[str, str]:
        if not self.api_key:
            raise ValueError("ANTHROPIC_API_KEY is not set.")
        return {
            "content-type": "application/json",
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
        }

    @staticmethod
    def _encode_pdf(pdf_bytes: bytes) -> str:
        return base64.b64encode(pdf_bytes).decode("utf-8")

    def extract_document(
        self, *, pdf_bytes: bytes, system_prompt: str, user_prompt: str
    ) -> dict[str, Any]:
        print(f"[anthropic] Sending request — model={self.model}, pdf={len(pdf_bytes) / 1024:.1f} KB, timeout={self.timeout_seconds}s")
        payload = {
            "model": self.model,
            "max_tokens": 8000,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": self._encode_pdf(pdf_bytes),
                            },
                        },
                        {
                            "type": "text",
                            "text": user_prompt,
                        },
                    ],
                }
            ],
            "system": system_prompt,
        }

        try:
            with httpx.Client(timeout=self.timeout_seconds) as client:
                response = client.post(self.BASE_URL, headers=self._build_headers(), json=payload)
                print(f"[anthropic] Response status: {response.status_code}")
                response.raise_for_status()
        except httpx.TimeoutException:
            print(f"[anthropic] ERROR: Request timed out after {self.timeout_seconds}s")
            raise
        except httpx.HTTPStatusError as exc:
            print(f"[anthropic] ERROR: HTTP {exc.response.status_code} — {exc.response.text[:300]}")
            raise
        except httpx.RequestError as exc:
            print(f"[anthropic] ERROR: Network error — {exc}")
            raise

        try:
            body = response.json()
        except Exception as exc:
            print(f"[anthropic] ERROR: Failed to parse response body as JSON — {exc}")
            raise ValueError(f"Anthropic returned non-JSON response: {exc}") from exc

        usage = body.get("usage", {})
        print(f"[anthropic] Tokens — input: {usage.get('input_tokens')}, output: {usage.get('output_tokens')}")

        text_chunks: list[str] = []
        for block in body.get("content", []):
            if block.get("type") == "text":
                text_chunks.append(block.get("text", ""))

        raw_text = "\n".join(text_chunks)
        print(f"[anthropic] Raw response length: {len(raw_text)} chars")

        if not raw_text.strip():
            print("[anthropic] ERROR: Model returned empty response")
            raise ValueError("Anthropic model returned an empty response.")

        parsed = extract_json_object(raw_text)
        parsed["_model_used"] = self.model
        print(f"[anthropic] JSON parsed successfully — top-level keys: {list(parsed.keys())}")
        return parsed


def make_primary_client() -> AnthropicClient:
    return AnthropicClient(
        api_key=settings.anthropic_api_key,
        model=settings.anthropic_model,
    )


def make_fallback_client() -> AnthropicClient:
    return AnthropicClient(
        api_key=settings.anthropic_api_key,
        model=settings.anthropic_fallback_model,
    )
