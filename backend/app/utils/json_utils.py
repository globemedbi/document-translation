from __future__ import annotations

import json
import re
from typing import Any


JSON_BLOCK_RE = re.compile(r"```json\s*(\{.*?\})\s*```", re.DOTALL)
OBJECT_RE = re.compile(r"(\{.*\})", re.DOTALL)


class JsonExtractionError(ValueError):
    pass


def extract_json_object(text: str) -> dict[str, Any]:
    text = text.strip()

    try:
        result = json.loads(text)
        print("[json_utils] Parsed JSON directly")
        return result
    except json.JSONDecodeError:
        pass

    block_match = JSON_BLOCK_RE.search(text)
    if block_match:
        try:
            result = json.loads(block_match.group(1))
            print("[json_utils] Extracted JSON from ```json block")
            return result
        except json.JSONDecodeError as exc:
            print(f"[json_utils] ERROR: Found ```json block but failed to parse it — {exc}")

    object_match = OBJECT_RE.search(text)
    if object_match:
        try:
            result = json.loads(object_match.group(1))
            print("[json_utils] Extracted JSON via regex fallback")
            return result
        except json.JSONDecodeError as exc:
            print(f"[json_utils] ERROR: Regex matched a JSON-like block but failed to parse it — {exc}")

    print("[json_utils] ERROR: Could not extract JSON from response")
    raise JsonExtractionError("Could not extract a valid JSON object from the model response.")
