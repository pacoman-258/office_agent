from __future__ import annotations

import io
import tempfile
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

from office_agent.cli import main
from office_agent.errors import RenderError
from office_agent.renderer import RenderResult
from office_agent.schema import PresentationSpec
from office_agent.services import normalize_download_filename


SPEC = PresentationSpec.model_validate(
    {
        "title": "Demo",
        "theme": {"preset": "executive"},
        "slides": [
            {"type": "title", "part": "opening", "title": "Intro", "subtitle": "Subtitle"},
            {"type": "summary", "part": "closing", "title": "Agenda", "key_points": ["A", "B"], "next_steps": ["Ship"]},
        ],
    }
)


class CLITests(TestCase):
    def test_normalize_download_filename_replaces_invalid_windows_characters(self) -> None:
        self.assertEqual(normalize_download_filename("演示:/文稿?.pptx"), "文稿_.pptx")

    def test_generate_writes_presentation_and_debug_spec(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            out_path = Path(temp_dir) / "demo.pptx"
            with patch("office_agent.cli.generate_spec_from_prompt", return_value=SPEC):
                with patch("office_agent.cli.render_presentation", return_value=RenderResult(path=out_path)):
                    with patch.dict("os.environ", {"OPENAI_API_KEY": "secret"}, clear=False):
                        code = main(
                            [
                                "generate",
                                "--prompt",
                                "Create a simple deck",
                                "--out",
                                str(out_path),
                                "--theme",
                                "editorial",
                                "--debug-spec",
                            ]
                        )
            self.assertEqual(code, 0)
            self.assertTrue(out_path.with_suffix(".spec.json").exists())

    def test_generate_fails_without_api_key_for_openai(self) -> None:
        stderr = io.StringIO()
        with patch("sys.stderr", stderr):
            with patch.dict("os.environ", {}, clear=True):
                code = main(["generate", "--prompt", "x", "--out", "out.pptx", "--provider", "openai"])
        self.assertEqual(code, 1)
        self.assertIn("OPENAI_API_KEY", stderr.getvalue())

    def test_generate_fails_when_rendering_raises(self) -> None:
        stderr = io.StringIO()
        with patch("office_agent.cli.generate_spec_from_prompt", return_value=SPEC):
            with patch("office_agent.cli.render_presentation", side_effect=RenderError("boom")):
                with patch.dict("os.environ", {"OPENAI_API_KEY": "secret"}, clear=False):
                    with patch("sys.stderr", stderr):
                        code = main(["generate", "--prompt", "x", "--out", "out.pptx", "--provider", "openai"])
        self.assertEqual(code, 1)
        self.assertIn("boom", stderr.getvalue())
