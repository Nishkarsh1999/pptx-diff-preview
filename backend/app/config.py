"""Runtime configuration, read from environment variables."""

import os
from dataclasses import dataclass


def _bool(name: str, default: str = "false") -> bool:
    return os.getenv(name, default).strip().lower() in ("1", "true", "yes", "on")


@dataclass(frozen=True)
class Settings:
    max_upload_mb: int = int(os.getenv("MAX_UPLOAD_MB", "100"))
    libreoffice_timeout: int = int(os.getenv("LIBREOFFICE_TIMEOUT_SECONDS", "120"))
    render_dpi: int = int(os.getenv("RENDER_DPI", "150"))
    match_threshold: float = float(os.getenv("MATCH_THRESHOLD", "0.55"))

    enable_ai_summary: bool = _bool("ENABLE_AI_SUMMARY")
    ai_provider: str = os.getenv("AI_PROVIDER", "anthropic").strip().lower()  # anthropic | openai
    ai_model: str = os.getenv("AI_MODEL", "claude-sonnet-4-6")
    anthropic_api_key: str = os.getenv("ANTHROPIC_API_KEY", "")
    openai_api_key: str = os.getenv("OPENAI_API_KEY", "")

    @property
    def ai_ready(self) -> bool:
        if not self.enable_ai_summary:
            return False
        if self.ai_provider == "anthropic":
            return bool(self.anthropic_api_key)
        if self.ai_provider == "openai":
            return bool(self.openai_api_key)
        return False


settings = Settings()
