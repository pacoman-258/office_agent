from __future__ import annotations

import base64
import tempfile
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

import requests
from pptx import Presentation
from pptx.util import Inches

from office_agent.renderer import PresentationRenderer
from office_agent.schema import PresentationSpec


PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnJyoAAAAAASUVORK5CYII="
)


class RendererTests(TestCase):
    def test_renderer_creates_pptx_with_supported_new_slide_types(self) -> None:
        spec = PresentationSpec.model_validate(
            {
                "title": "Demo",
                "theme": {
                    "preset": "editorial",
                    "custom": {"primary_color": "#224466", "cover_style": "minimal"},
                },
                "slides": [
                    {"type": "title", "part": "opening", "title": "Intro", "subtitle": "Subtitle"},
                    {
                        "type": "timeline",
                        "part": "content",
                        "title": "Roadmap",
                        "events": [
                            {"label": "Q1", "title": "Plan", "detail": "Scope"},
                            {"label": "Q2", "title": "Ship", "detail": "Launch"},
                        ],
                    },
                    {"type": "quote", "part": "content", "title": "Voice", "quote": "Stay close to users.", "attribution": "PM"},
                    {
                        "type": "comparison",
                        "part": "content",
                        "title": "Compare",
                        "left": {"title": "Option A", "bullets": ["Fast"]},
                        "right": {"title": "Option B", "bullets": ["Safe"]},
                    },
                    {"type": "summary", "part": "closing", "title": "Summary", "key_points": ["One", "Two"], "next_steps": ["Act"]},
                    {"type": "table", "part": "content", "title": "Budget", "headers": ["Item", "Owner"], "rows": [["Design", "Alice"], ["Build", "Bob"]]},
                ],
            }
        )
        renderer = PresentationRenderer()
        with tempfile.TemporaryDirectory() as temp_dir:
            out_path = Path(temp_dir) / "demo.pptx"
            result = renderer.render(spec, out_path)
            self.assertTrue(result.path.exists())
            presentation = Presentation(str(result.path))
            self.assertEqual(len(presentation.slides), 6)

    def test_renderer_downgrades_missing_image_url_to_text_slide(self) -> None:
        spec = PresentationSpec.model_validate(
            {
                "title": "Demo",
                "slides": [
                    {
                        "type": "image",
                        "part": "content",
                        "title": "Image",
                        "image": "https://example.com/missing.png",
                        "caption": "Fallback",
                        "bullets": ["Note"],
                    }
                ],
            }
        )
        renderer = PresentationRenderer()
        with tempfile.TemporaryDirectory() as temp_dir:
            out_path = Path(temp_dir) / "image.pptx"
            with patch("office_agent.renderer.requests.get", side_effect=requests.RequestException("boom")):
                result = renderer.render(spec, out_path)
            self.assertTrue(result.warnings)
            presentation = Presentation(str(result.path))
            self.assertEqual(len(presentation.slides), 1)

    def test_renderer_embeds_local_image(self) -> None:
        spec = PresentationSpec.model_validate(
            {
                "title": "Demo",
                "slides": [
                    {
                        "type": "image",
                        "part": "content",
                        "title": "Image",
                        "image": "placeholder",
                        "caption": "Local",
                        "bullets": ["Point"],
                    }
                ],
            }
        )
        renderer = PresentationRenderer()
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "pixel.png"
            image_path.write_bytes(PNG_BYTES)
            spec.slides[0].image = str(image_path)
            out_path = Path(temp_dir) / "local.pptx"
            result = renderer.render(spec, out_path)
            self.assertFalse(result.warnings)
            self.assertTrue(result.path.exists())

    def test_renderer_uses_template_slide_for_selected_part(self) -> None:
        spec = PresentationSpec.model_validate(
            {
                "title": "Demo",
                "template": {"opening": 0, "agenda": 1, "content": 1, "closing": 1},
                "slides": [{"type": "title", "part": "opening", "title": "Generated", "subtitle": "From template"}],
            }
        )
        renderer = PresentationRenderer()
        with tempfile.TemporaryDirectory() as temp_dir:
            template_path = Path(temp_dir) / "template.pptx"
            template = Presentation()
            slide = template.slides.add_slide(template.slide_layouts[6])
            title_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
            title_box.name = "oa:title"
            subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(1))
            subtitle_box.name = "oa:subtitle"
            template.save(str(template_path))

            out_path = Path(temp_dir) / "templated-output.pptx"
            result = renderer.render(spec, out_path, template_path=template_path)

            presentation = Presentation(str(result.path))
            self.assertEqual(len(presentation.slides), 1)
            texts = [shape.text for shape in presentation.slides[0].shapes if getattr(shape, "has_text_frame", False)]
            self.assertIn("Generated", texts)
            self.assertIn("From template", texts)

    def test_renderer_falls_back_when_template_placeholders_are_missing(self) -> None:
        spec = PresentationSpec.model_validate(
            {
                "title": "Demo",
                "template": {"opening": 0, "agenda": 0, "content": 0, "closing": 0},
                "slides": [{"type": "summary", "part": "closing", "title": "Wrap", "key_points": ["One"], "next_steps": ["Act"]}],
            }
        )
        renderer = PresentationRenderer()
        with tempfile.TemporaryDirectory() as temp_dir:
            template_path = Path(temp_dir) / "template.pptx"
            template = Presentation()
            slide = template.slides.add_slide(template.slide_layouts[6])
            title_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
            title_box.name = "oa:title"
            template.save(str(template_path))

            out_path = Path(temp_dir) / "fallback-output.pptx"
            result = renderer.render(spec, out_path, template_path=template_path)

            self.assertTrue(result.warnings)
            presentation = Presentation(str(result.path))
            self.assertEqual(len(presentation.slides), 1)
