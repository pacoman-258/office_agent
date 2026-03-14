from __future__ import annotations

from dataclasses import dataclass

import requests

from office_agent.config import AppConfig
from office_agent.errors import ProviderError


@dataclass
class BaseProvider:
    config: AppConfig
    session: requests.Session | None = None

    @property
    def http(self) -> requests.Session:
        return self.session or requests.Session()


class OpenAICompatibleProvider(BaseProvider):
    def generate_text(self, messages: list[dict[str, str]], model: str) -> str:
        url = f"{self.config.openai_base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.config.openai_api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": model,
            "messages": messages,
            "response_format": {"type": "json_object"},
        }
        response = self.http.post(
            url,
            headers=headers,
            json=payload,
            timeout=self.config.timeout_seconds,
        )
        try:
            response.raise_for_status()
            data = response.json()
            return data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, ValueError, requests.RequestException) as exc:
            raise ProviderError(f"OpenAI-compatible request failed: {exc}") from exc


class OllamaProvider(BaseProvider):
    def generate_text(self, messages: list[dict[str, str]], model: str) -> str:
        url = f"{self.config.ollama_base_url}/api/chat"
        payload = {
            "model": model,
            "messages": messages,
            "stream": False,
            "format": "json",
        }
        response = self.http.post(url, json=payload, timeout=self.config.timeout_seconds)
        try:
            response.raise_for_status()
            data = response.json()
            return data["message"]["content"]
        except (KeyError, ValueError, requests.RequestException) as exc:
            raise ProviderError(f"Ollama request failed: {exc}") from exc
