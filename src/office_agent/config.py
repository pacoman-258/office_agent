from __future__ import annotations

import os
from dataclasses import dataclass, field

from office_agent.errors import ConfigError
from office_agent.schema import ThemeSpec


DEFAULT_PROVIDER = "openai"
DEFAULT_MODEL = "qwen3:8b"
DEFAULT_THEME_PRESET = "default"
DEFAULT_OPENAI_BASE_URL = "https://api.openai.com/v1"
DEFAULT_OLLAMA_BASE_URL = "http://localhost:11434"


@dataclass(frozen=True)
class AppConfig:
    provider: str = DEFAULT_PROVIDER
    model: str = DEFAULT_MODEL
    theme: ThemeSpec = field(default_factory=ThemeSpec)
    openai_api_key: str | None = None
    openai_base_url: str = DEFAULT_OPENAI_BASE_URL
    ollama_base_url: str = DEFAULT_OLLAMA_BASE_URL
    timeout_seconds: float = 60.0

    @classmethod
    def from_env(cls) -> "AppConfig":
        return cls(
            provider=os.getenv("OFFICE_AGENT_PROVIDER", DEFAULT_PROVIDER),
            model=os.getenv("OFFICE_AGENT_MODEL", DEFAULT_MODEL),
            theme=ThemeSpec(preset=DEFAULT_THEME_PRESET),
            openai_api_key=os.getenv("OPENAI_API_KEY"),
            openai_base_url=os.getenv("OPENAI_BASE_URL", DEFAULT_OPENAI_BASE_URL).rstrip("/"),
            ollama_base_url=os.getenv("OLLAMA_BASE_URL", DEFAULT_OLLAMA_BASE_URL).rstrip("/"),
        )

    def with_overrides(
        self,
        *,
        provider: str | None = None,
        model: str | None = None,
        theme: ThemeSpec | None = None,
        openai_api_key: str | None = None,
        openai_base_url: str | None = None,
        ollama_base_url: str | None = None,
    ) -> "AppConfig":
        return AppConfig(
            provider=provider or self.provider,
            model=model or self.model,
            theme=theme or self.theme,
            openai_api_key=openai_api_key if openai_api_key is not None else self.openai_api_key,
            openai_base_url=(openai_base_url or self.openai_base_url).rstrip("/"),
            ollama_base_url=(ollama_base_url or self.ollama_base_url).rstrip("/"),
            timeout_seconds=self.timeout_seconds,
        )

    def validate(self) -> None:
        if self.provider not in {"openai", "ollama"}:
            raise ConfigError(f"Unsupported provider: {self.provider}")
        if self.provider == "openai" and not self.openai_api_key:
            raise ConfigError("OPENAI_API_KEY is required when provider is openai.")
