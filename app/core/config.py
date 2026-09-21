# -*- coding: utf-8 -*-
"""Application configuration settings.

Uses Pydantic BaseSettings to load from environment variables and .env file.
Implements lazy initialization so importing the module doesn't fail if env vars
are missing (useful for tests and CLI tooling).
"""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings loaded from environment variables or .env file."""

    # Gemini
    GEMINI_API_KEY: str = ""
    GEMINI_MODEL: str = "gemini-3.6-flash"

    # Supabase
    SUPABASE_URL: str = ""
    SUPABASE_KEY: str = ""
    SUPABASE_BUCKET: str = "comparativos"

    # Paths
    ASSETS_DIR: Path = Path(__file__).resolve().parent.parent.parent / "assets"
    TEMPLATES_DIR: Path = Path(__file__).resolve().parent.parent.parent / "assets"

    # Limits
    MAX_FILE_SIZE_MB: int = 50

    # Logging
    LOG_LEVEL: str = "INFO"

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    def validate_required(self) -> None:
        """Validate that all required settings are present.

        Raises:
            ValueError: If any required setting is missing.
        """
        missing = []
        if not self.GEMINI_API_KEY:
            missing.append("GEMINI_API_KEY")
        if not self.SUPABASE_URL:
            missing.append("SUPABASE_URL")
        if not self.SUPABASE_KEY:
            missing.append("SUPABASE_KEY")
        if missing:
            raise ValueError(
                f"Missing required environment variables: {', '.join(missing)}. "
                f"Set them in .env or as environment variables."
            )


@lru_cache()
def get_settings() -> Settings:
    """Create and return a cached application settings instance."""
    return Settings()


settings = get_settings()
