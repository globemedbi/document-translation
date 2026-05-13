from __future__ import annotations

import os
import sys
from enum import Enum
from pathlib import Path

from dotenv import load_dotenv
from loguru import logger

load_dotenv()


class SolutionMode(str, Enum):
    VISION = "VISION"
    CODING = "CODING"
    TEXT = "TEXT"    # Direct text-layer replacement for selectable-text PDFs
    AUTO = "AUTO"    # Auto-detect: TEXT if embedded text found, else VISION


class Settings:
    def __init__(self) -> None:
        self.anthropic_api_key: str = self._require("ANTHROPIC_API_KEY")
        self.anthropic_model: str = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-6")
        # Heavy model for extraction & coding agent (defaults to Opus for best quality)
        self.anthropic_heavy_model: str = os.getenv(
            "ANTHROPIC_HEAVY_MODEL", "claude-opus-4-7"
        )

        raw_mode = os.getenv("SOLUTION_MODE", "VISION").upper().strip()
        try:
            self.solution_mode = SolutionMode(raw_mode)
        except ValueError:
            raise ValueError(
                f"SOLUTION_MODE must be VISION, CODING, TEXT, or AUTO, got: {raw_mode!r}"
            )

        self.target_language: str = os.getenv("TARGET_LANGUAGE", "French")
        self.output_dir: Path = Path(os.getenv("OUTPUT_DIR", "output"))
        self.max_agent_iterations: int = int(os.getenv("MAX_AGENT_ITERATIONS", "3"))
        self.agent_workspace: Path = Path(os.getenv("AGENT_WORKSPACE", "workspace"))
        self.log_level: str = os.getenv("LOG_LEVEL", "INFO").upper()

    @staticmethod
    def _require(key: str) -> str:
        val = os.getenv(key, "").strip()
        if not val:
            raise ValueError(
                f"Missing required environment variable: {key}\n"
                f"Copy .env.example to .env and fill in the value."
            )
        return val

    def __repr__(self) -> str:
        return (
            f"Settings(mode={self.solution_mode}, lang={self.target_language!r}, "
            f"model={self.anthropic_model!r}, heavy={self.anthropic_heavy_model!r})"
        )


settings = Settings()
