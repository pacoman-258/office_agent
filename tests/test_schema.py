from unittest import TestCase

from pydantic import ValidationError

from office_agent.schema import PresentationSpec


class SchemaTests(TestCase):
    def test_accepts_all_supported_slide_types_theme_and_template_mapping(self) -> None:
        spec = PresentationSpec.model_validate(
            {
                "title": "Demo",
                "theme": {
                    "preset": "executive",
                    "custom": {
                        "primary_color": "#112233",
                        "accent_color": "#445566",
                        "background_color": "#F5F5F5",
                        "heading_font": "Georgia",
                        "body_font": "Microsoft YaHei",
                        "cover_style": "centered",
                    },
                },
                "template": {"opening": 0, "agenda": 1, "content": 2, "closing": 3},
                "slides": [
                    {"type": "title", "part": "opening", "title": "Intro", "subtitle": "Sub"},
                    {"type": "section", "part": "agenda", "title": "Part 1", "subtitle": "Overview"},
                    {"type": "bullets", "part": "content", "title": "Highlights", "bullets": ["A", "B"]},
                    {
                        "type": "two_column",
                        "part": "content",
                        "title": "Compare",
                        "left_title": "Left",
                        "left_bullets": ["L1"],
                        "right_title": "Right",
                        "right_bullets": ["R1"],
                    },
                    {
                        "type": "image",
                        "part": "content",
                        "title": "Image",
                        "image": "https://example.com/image.png",
                        "caption": "Caption",
                        "bullets": ["Point"],
                    },
                    {
                        "type": "timeline",
                        "part": "content",
                        "title": "Roadmap",
                        "events": [
                            {"label": "Q1", "title": "Plan", "detail": "Scope"},
                            {"label": "Q2", "title": "Launch", "detail": "Ship"},
                        ],
                    },
                    {
                        "type": "quote",
                        "part": "content",
                        "title": "Voice",
                        "quote": "Stay close to users",
                        "attribution": "PM",
                        "source": "Kickoff",
                    },
                    {
                        "type": "comparison",
                        "part": "content",
                        "title": "Compare",
                        "left": {"title": "Option A", "bullets": ["Fast"]},
                        "right": {"title": "Option B", "bullets": ["Cheap"]},
                    },
                    {
                        "type": "summary",
                        "part": "closing",
                        "title": "Wrap Up",
                        "key_points": ["One", "Two"],
                        "next_steps": ["Act"],
                    },
                    {
                        "type": "table",
                        "part": "content",
                        "title": "Budget",
                        "headers": ["Item", "Owner"],
                        "rows": [["Design", "Alice"], ["Build", "Bob"]],
                    },
                ],
            }
        )
        self.assertEqual(spec.theme.preset, "executive")
        self.assertEqual(spec.template.content, 2)
        self.assertEqual(len(spec.slides), 10)

    def test_rejects_invalid_theme_color(self) -> None:
        with self.assertRaises(ValidationError):
            PresentationSpec.model_validate(
                {
                    "title": "Broken",
                    "theme": {"preset": "default", "custom": {"primary_color": "blue"}},
                    "slides": [{"type": "title", "part": "opening", "title": "Intro"}],
                }
            )

    def test_rejects_invalid_table_shape(self) -> None:
        with self.assertRaises(ValidationError):
            PresentationSpec.model_validate(
                {
                    "title": "Broken",
                    "slides": [
                        {
                            "type": "table",
                            "part": "content",
                            "title": "Table",
                            "headers": ["A", "B"],
                            "rows": [["Only one"]],
                        }
                    ],
                }
            )

    def test_rejects_invalid_slide_type(self) -> None:
        with self.assertRaises(ValidationError):
            PresentationSpec.model_validate(
                {
                    "title": "Broken",
                    "slides": [{"type": "chart", "part": "content", "title": "Unsupported"}],
                }
            )

    def test_rejects_missing_part(self) -> None:
        with self.assertRaises(ValidationError):
            PresentationSpec.model_validate(
                {
                    "title": "Broken",
                    "slides": [{"type": "title", "title": "No part"}],
                }
            )
