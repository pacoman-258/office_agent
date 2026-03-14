from __future__ import annotations

import json
from unittest import TestCase
from unittest.mock import Mock, patch
from urllib.parse import unquote

from fastapi.testclient import TestClient

from office_agent.api.app import create_app
from office_agent.api.models import TemplatePreviewResponse, TemplatePreviewSlide
from office_agent.schema import PresentationSpec


SPEC = PresentationSpec.model_validate(
    {
        "title": "Demo",
        "theme": {"preset": "executive", "custom": {"primary_color": "#112233"}},
        "template": {"opening": 0, "agenda": 1, "content": 2, "closing": 3},
        "slides": [
            {"type": "title", "part": "opening", "title": "Intro", "subtitle": "Subtitle"},
            {"type": "summary", "part": "closing", "title": "Agenda", "key_points": ["A", "B"], "next_steps": ["Ship"]},
        ],
    }
)


class ApiTests(TestCase):
    def setUp(self) -> None:
        self.client = TestClient(create_app())

    def test_health_returns_supported_providers_and_themes(self) -> None:
        response = self.client.get("/api/health")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["providers"], ["openai", "ollama"])
        self.assertEqual(response.json()["themes"], ["default", "executive", "editorial"])

    def test_specs_requires_api_key_for_openai(self) -> None:
        response = self.client.post(
            "/api/specs",
            json={
                "prompt": "Create a short deck",
                "provider": "openai",
                "model": "gpt-4o-mini",
                "theme": {"preset": "default"},
            },
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("OPENAI_API_KEY", response.json()["detail"])

    def test_specs_maps_runtime_provider_config_and_template_mapping(self) -> None:
        with patch("office_agent.api.app.generate_spec_from_prompt", return_value=SPEC) as mock_generate:
            response = self.client.post(
                "/api/specs",
                json={
                    "prompt": "Create a short deck",
                    "provider": "openai",
                    "model": "gpt-4o-mini",
                    "theme": {"preset": "editorial", "custom": {"cover_style": "minimal"}},
                    "templateMapping": {"opening": 0, "agenda": 1, "content": 2, "closing": 3},
                    "apiKey": "secret",
                    "openaiBaseUrl": "https://example.com/v1",
                    "ollamaBaseUrl": "http://localhost:11434",
                },
            )
        self.assertEqual(response.status_code, 200)
        _, config = mock_generate.call_args.args[:2]
        template_mapping = mock_generate.call_args.kwargs["template_mapping"]
        self.assertEqual(config.provider, "openai")
        self.assertEqual(config.openai_api_key, "secret")
        self.assertEqual(config.openai_base_url, "https://example.com/v1")
        self.assertEqual(config.theme.preset, "editorial")
        self.assertEqual(template_mapping.content, 2)

    def test_template_preview_returns_slides(self) -> None:
        preview = TemplatePreviewResponse(
            slides=[
                TemplatePreviewSlide(index=0, thumbnailDataUrl="data:image/png;base64,aaa", titleText="Cover", placeholderRoles=["title", "subtitle"])
            ]
        )
        with patch("office_agent.api.app.preview_template_artifact", return_value=preview):
            response = self.client.post(
                "/api/templates/preview",
                files={"template": ("template.pptx", b"pptx-bytes", "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
            )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["slides"][0]["index"], 0)
        self.assertEqual(response.json()["slides"][0]["placeholderRoles"], ["title", "subtitle"])

    def test_presentations_returns_binary_pptx(self) -> None:
        with patch("office_agent.api.app.render_presentation_artifact") as mock_render:
            mock_render.return_value = Mock(filename="deck.pptx", content=b"pptx-bytes", warnings=[])
            response = self.client.post("/api/presentations", json={"filename": "deck", "spec": SPEC.model_dump(mode="json")})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["content-type"], "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        self.assertEqual(response.content, b"pptx-bytes")
        self.assertIn("filename*=UTF-8''deck.pptx", response.headers["content-disposition"])

    def test_presentations_rejects_invalid_spec(self) -> None:
        response = self.client.post(
            "/api/presentations",
            json={"filename": "deck", "spec": {"title": "Broken", "slides": [{"type": "chart", "part": "content"}]}},
        )
        self.assertEqual(response.status_code, 422)

    def test_presentations_encode_non_ascii_warnings_in_header(self) -> None:
        with patch("office_agent.api.app.render_presentation_artifact") as mock_render:
            mock_render.return_value = Mock(filename="演示文稿.pptx", content=b"pptx-bytes", warnings=["Image path not found: 图片1.png"])
            response = self.client.post("/api/presentations", json={"filename": "演示文稿", "spec": SPEC.model_dump(mode="json")})
        self.assertEqual(response.status_code, 200)
        warnings = json.loads(unquote(response.headers["x-office-agent-warnings"]))
        self.assertEqual(warnings, ["Image path not found: 图片1.png"])

    def test_presentations_accepts_multipart_template_upload(self) -> None:
        with patch("office_agent.api.app.render_presentation_artifact") as mock_render:
            mock_render.return_value = Mock(filename="deck.pptx", content=b"pptx-bytes", warnings=[])
            response = self.client.post(
                "/api/presentations",
                data={"payload": json.dumps({"filename": "deck", "spec": SPEC.model_dump(mode="json")})},
                files={
                    "template": (
                        "template.pptx",
                        b"template-bytes",
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
                },
            )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(mock_render.call_args.kwargs["template_bytes"], b"template-bytes")
        self.assertEqual(mock_render.call_args.kwargs["template_filename"], "template.pptx")
