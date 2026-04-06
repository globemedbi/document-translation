import os
from pathlib import Path

from dotenv import load_dotenv
from pydantic_settings import BaseSettings, SettingsConfigDict

load_dotenv()
key: str | None = os.environ.get("ANTHROPIC_API_KEY")


class Settings(BaseSettings):
    anthropic_api_key: str | None = key
    anthropic_model: str = "claude-sonnet-4-6"
    anthropic_fallback_model: str = "claude-opus-4-6"
    backend_base_url: str = "http://localhost:8000"
    storage_dir: str = "storage"
    max_file_size_mb: int = 32
    max_pdf_pages: int = 100
    target_language: str = "English"

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    @property
    def storage_path(self) -> Path:
        return Path(self.storage_dir)


settings = Settings()
