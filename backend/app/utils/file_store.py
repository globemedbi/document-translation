from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any


class JobStore:
    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def job_dir(self, job_id: str) -> Path:
        path = self.base_dir / job_id
        path.mkdir(parents=True, exist_ok=True)
        return path

    def save_upload(self, job_id: str, filename: str, content: bytes) -> Path:
        target = self.job_dir(job_id) / filename
        try:
            target.write_bytes(content)
            print(f"[file_store] Saved upload: {target} ({len(content) / 1024:.1f} KB)")
        except OSError as exc:
            print(f"[file_store] ERROR: Failed to save upload {target} — {exc}")
            raise
        return target

    def save_json(self, job_id: str, filename: str, payload: dict[str, Any]) -> Path:
        target = self.job_dir(job_id) / filename
        try:
            target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"[file_store] Saved JSON: {target}")
        except OSError as exc:
            print(f"[file_store] ERROR: Failed to save JSON {target} — {exc}")
            raise
        return target

    def save_file(self, job_id: str, filename: str, source_path: Path) -> Path:
        target = self.job_dir(job_id) / filename
        try:
            shutil.copy2(source_path, target)
            print(f"[file_store] Copied file to: {target}")
        except OSError as exc:
            print(f"[file_store] ERROR: Failed to copy file {source_path} -> {target} — {exc}")
            raise
        return target

    def get_file(self, job_id: str, filename: str) -> Path:
        path = self.base_dir / job_id / filename
        if not path.exists():
            print(f"[file_store] ERROR: File not found — {path}")
            raise FileNotFoundError(f"No file at {path}")
        return path
