from __future__ import annotations

import json
from typing import Any

import requests

from office_agent.errors import OfficeAutomationError
from office_agent.office.models import FinalizeConfig, ReviewContext, ReviewDeckResult
from office_agent.office.prompts import build_review_messages


class VisualReviewer:
    def __init__(self, config: FinalizeConfig, session: requests.Session | None = None) -> None:
        self.config = config
        self.session = session or requests.Session()

    def review(self, context: ReviewContext) -> ReviewDeckResult:
        if self.config.provider != "openai":
            raise OfficeAutomationError(f"Unsupported finalize provider: {self.config.provider}")
        if not self.config.api_key:
            raise OfficeAutomationError("Visual finalization requires an OpenAI-compatible API key.")
        if not self.config.model:
            raise OfficeAutomationError("Visual finalization requires a model name.")

        previous_error: str | None = None
        for _ in range(2):
            messages = build_review_messages(context)
            if previous_error:
                messages.append(
                    {
                        "role": "user",
                        "content": f"Your previous response was invalid. Fix the JSON only.\nError: {previous_error}",
                    }
                )
            raw = self._send_request(messages)
            try:
                payload = _extract_json_payload(raw)
                return ReviewDeckResult.model_validate(payload)
            except Exception as exc:  # pragma: no cover - simple retry guard
                previous_error = str(exc)
        raise OfficeAutomationError(f"Visual reviewer did not return valid JSON: {previous_error}")

    def _send_request(self, messages: list[dict[str, object]]) -> str:
        url = f"{self.config.base_url.rstrip('/')}/chat/completions"
        response = self.session.post(
            url,
            headers={
                "Authorization": f"Bearer {self.config.api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": self.config.model,
                "messages": messages,
                "response_format": {"type": "json_object"},
            },
            timeout=90,
        )
        try:
            response.raise_for_status()
            data = response.json()
            return data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, ValueError, requests.RequestException) as exc:
            raise OfficeAutomationError(f"Visual reviewer request failed: {exc}") from exc


def _extract_json_payload(text: str) -> dict[str, Any]:
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
        raise OfficeAutomationError("Visual reviewer response did not contain a JSON object.")
    try:
        return json.loads(raw[start : end + 1])
    except json.JSONDecodeError as exc:
        raise OfficeAutomationError(f"Visual reviewer returned invalid JSON: {exc}") from exc
