from unittest import TestCase
from unittest.mock import Mock

from office_agent.config import AppConfig
from office_agent.errors import SpecGenerationError
from office_agent.llm.service import extract_json_payload, generate_presentation_spec
from office_agent.schema import TemplateSelectionSpec, ThemeSpec


class LLMServiceTests(TestCase):
    def test_extract_json_payload_supports_fenced_blocks(self) -> None:
        payload = extract_json_payload(
            """```json
{"title":"Demo","theme":{"preset":"default"},"template":{"opening":0,"agenda":1,"content":2,"closing":3},"slides":[{"type":"title","part":"opening","title":"A"}]}
```"""
        )
        self.assertEqual(payload["title"], "Demo")

    def test_generate_presentation_spec_retries_once_and_overrides_theme_and_template(self) -> None:
        session = Mock()
        response_bad = Mock()
        response_bad.raise_for_status.return_value = None
        response_bad.json.return_value = {"message": {"content": "not json"}}
        response_good = Mock()
        response_good.raise_for_status.return_value = None
        response_good.json.return_value = {
            "message": {
                "content": (
                    '{"title":"Demo","theme":{"preset":"default"},"template":{"opening":0,"agenda":1,"content":2,"closing":3},'
                    '"slides":[{"type":"title","part":"opening","title":"A"},{"type":"summary","part":"closing","title":"B","key_points":["x"],"next_steps":["ship"]}]}'
                )
            }
        }
        session.post.side_effect = [response_bad, response_good]

        config = AppConfig(provider="ollama", model="qwen2.5", theme=ThemeSpec(preset="editorial"))
        mapping = TemplateSelectionSpec(opening=0, agenda=1, content=2, closing=3)
        spec = generate_presentation_spec("Create a short deck", config.model, config, template_mapping=mapping, session=session)
        self.assertEqual(len(spec.slides), 2)
        self.assertEqual(session.post.call_count, 2)
        self.assertEqual(spec.theme.preset, "editorial")
        self.assertEqual(spec.template, mapping)

    def test_generate_presentation_spec_fails_after_retry(self) -> None:
        session = Mock()
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {"message": {"content": "still wrong"}}
        session.post.side_effect = [response, response]

        config = AppConfig(provider="ollama", model="qwen2.5", theme=ThemeSpec(preset="default"))
        with self.assertRaises(SpecGenerationError):
            generate_presentation_spec("Create a short deck", config.model, config, session=session)

    def test_openai_provider_uses_expected_shape(self) -> None:
        session = Mock()
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "choices": [
                {
                    "message": {
                        "content": (
                            '{"title":"Demo","theme":{"preset":"default"},"slides":[{"type":"title","part":"opening","title":"A"}]}'
                        )
                    }
                }
            ]
        }
        session.post.return_value = response

        config = AppConfig(provider="openai", model="gpt-4o-mini", openai_api_key="secret", theme=ThemeSpec())
        spec = generate_presentation_spec("Create slides", config.model, config, session=session)
        self.assertEqual(spec.title, "Demo")
        _, kwargs = session.post.call_args
        self.assertIn("response_format", kwargs["json"])
