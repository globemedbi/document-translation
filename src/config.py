from __future__ import annotations

import os
from enum import Enum
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


class SolutionMode(str, Enum):
    VISION = "VISION"
    CODING = "CODING"
    TEXT = "TEXT"
    AUTO = "AUTO"


class Settings:
    def __init__(self) -> None:
        self.llm_provider: str = os.getenv("LLM_PROVIDER", "anthropic").lower().strip()

        self.anthropic_api_key: str = os.getenv("ANTHROPIC_API_KEY", "")
        self.anthropic_model: str = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-6")
        self.anthropic_heavy_model: str = os.getenv("ANTHROPIC_HEAVY_MODEL", "claude-opus-4-7")

        self.openai_api_key: str = os.getenv("OPENAI_API_KEY", "")
        self.openai_model: str = os.getenv("OPENAI_MODEL", "gpt-5.5")
        self.openai_heavy_model: str = os.getenv("OPENAI_HEAVY_MODEL", "o4-mini")
        self.openai_coding_model: str = os.getenv("OPENAI_CODING_MODEL", "o3")

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

    @property
    def model(self) -> str:
        return self.openai_model if self.llm_provider == "openai" else self.anthropic_model

    @property
    def heavy_model(self) -> str:
        return (
            self.openai_heavy_model if self.llm_provider == "openai" else self.anthropic_heavy_model
        )

    @property
    def coding_model(self) -> str:
        return self.openai_coding_model if self.llm_provider == "openai" else self.anthropic_model

    def __repr__(self) -> str:
        return (
            f"Settings(provider={self.llm_provider}, mode={self.solution_mode}, "
            f"lang={self.target_language!r}, model={self.model!r})"
        )


settings = Settings()


def get_client():
    """Return the configured LLM client (Anthropic or OpenAI)."""
    if settings.llm_provider == "openai":
        import openai

        return openai.OpenAI(api_key=settings.openai_api_key)
    import anthropic

    return anthropic.Anthropic(api_key=settings.anthropic_api_key)
