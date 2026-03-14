from __future__ import annotations

import json
from typing import Any

import requests
from pydantic import ValidationError

from office_agent.config import AppConfig
from office_agent.errors import SpecGenerationError
from office_agent.llm.prompts import build_messages
from office_agent.llm.providers import OllamaProvider, OpenAICompatibleProvider
from office_agent.schema import PresentationSpec, TemplateSelectionSpec


def create_provider(config: AppConfig, session: requests.Session | None = None):
    if config.provider == "openai":
        return OpenAICompatibleProvider(config=config, session=session)
    if config.provider == "ollama":
        return OllamaProvider(config=config, session=session)
    raise SpecGenerationError(f"Unsupported provider: {config.provider}")


def extract_json_payload(text: str) -> dict[str, Any]:
    raw = text.strip()
    if raw.startswith("```"):
        lines = raw.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        raw = "\n".join(lines).strip()

    start = raw.find("{")
    end = raw.rfind("}")
    if start == -1 or end == -1 or end < start:
        raise SpecGenerationError("Provider response did not contain a JSON object.")
    try:
        return json.loads(raw[start : end + 1])
    except json.JSONDecodeError as exc:
        raise SpecGenerationError(f"Invalid JSON returned by provider: {exc}") from exc


def generate_presentation_spec(
    prompt: str,
    model: str,
    config: AppConfig,
    template_mapping: TemplateSelectionSpec | None = None,
    session: requests.Session | None = None,
) -> PresentationSpec:
    provider = create_provider(config, session=session)
    previous_error: str | None = None
    for _ in range(2):
        messages = build_messages(
            prompt,
            requested_theme=config.theme,
            template_mapping=template_mapping,
            previous_error=previous_error,
        )
        raw_text = provider.generate_text(messages=messages, model=model)
        try:
            payload = extract_json_payload(raw_text)
            spec = PresentationSpec.model_validate(payload)
            updates = {"theme": config.theme}
            if template_mapping is not None:
                updates["template"] = template_mapping
            return spec.model_copy(update=updates)
        except (SpecGenerationError, ValidationError) as exc:
            previous_error = str(exc)
    raise SpecGenerationError(f"Unable to generate a valid presentation spec: {previous_error}")
