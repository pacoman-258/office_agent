from unittest import TestCase
from unittest.mock import Mock

from office_agent.config import AppConfig
from office_agent.llm.service import generate_presentation_spec
from office_agent.schema import ThemeSpec


class SmokeTests(TestCase):
    def test_prompt_generates_three_to_five_slides(self) -> None:
        session = Mock()
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "message": {
                "content": (
                    '{"title":"AI Office Agent","theme":{"preset":"default"},"slides":['
                    '{"type":"title","part":"opening","title":"AI Office Agent","subtitle":"Phase One"},'
                    '{"type":"summary","part":"closing","title":"Core Capabilities","key_points":["Natural language input","Structured rendering","PPT output"],"next_steps":["Ship pilot"]},'
                    '{"type":"image","part":"content","title":"Use Cases","image":"https://example.com/demo.png","caption":"Reference image","bullets":["Scenario overview"]}'
                    "]}"
                )
            }
        }
        session.post.return_value = response

        config = AppConfig(provider="ollama", model="qwen2.5", theme=ThemeSpec(preset="executive"))
        spec = generate_presentation_spec("Generate a presentation about an AI office agent", config.model, config, session=session)
        self.assertGreaterEqual(len(spec.slides), 3)
        self.assertLessEqual(len(spec.slides), 5)
        slide_types = [slide.type for slide in spec.slides]
        self.assertIn("title", slide_types)
        self.assertIn("summary", slide_types)
        self.assertIn("image", slide_types)
        self.assertEqual(spec.theme.preset, "executive")
