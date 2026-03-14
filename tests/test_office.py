from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch

from office_agent.office.finalizer import finalize_presentation
from office_agent.office.models import (
    FinalizeConfig,
    ReviewDeckResult,
    SlideEditOperation,
    SlideIssue,
    SlideReviewInput,
    SlideReviewResult,
)
from office_agent.schema import PresentationSpec


SPEC = PresentationSpec.model_validate(
    {
        "title": "Demo",
        "slides": [{"type": "bullets", "part": "content", "title": "Intro", "bullets": ["A", "B"]}],
    }
)


class OfficeFinalizerTests(TestCase):
    def test_finalize_skips_when_disabled(self) -> None:
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "demo.pptx"
            path.write_bytes(b"ppt")
            result = finalize_presentation(path, spec=SPEC, config=FinalizeConfig(enabled=False))
        self.assertEqual(result.summary.status, "skipped")
        self.assertEqual(result.summary.rounds, [])

    def test_finalize_skips_when_visual_model_missing(self) -> None:
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "demo.pptx"
            path.write_bytes(b"ppt")
            result = finalize_presentation(path, spec=SPEC, config=FinalizeConfig(enabled=True, model=None, api_key="secret"))
        self.assertEqual(result.summary.status, "skipped")
        self.assertTrue(result.summary.warnings)

    def test_finalize_runs_review_loop_and_records_summary(self) -> None:
        review_result = ReviewDeckResult(
            slides=[
                SlideReviewResult(
                    slideIndex=0,
                    issues=[
                        SlideIssue(
                            severity="high",
                            reason="Text overflow",
                            operations=[SlideEditOperation(type="set_font_size", slideIndex=0, shapeId=3, fontSize=20)],
                        )
                    ],
                )
            ]
        )
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "demo.pptx"
            path.write_bytes(b"ppt")
            with patch(
                "office_agent.office.finalizer.export_review_inputs",
                return_value=[
                    SlideReviewInput(
                        slideIndex=0,
                        slideTitle="Intro",
                        slideType="bullets",
                        textSummary=["Intro", "A", "B"],
                        imageDataUrl="data:image/png;base64,aaa",
                        shapes=[],
                    )
                ],
            ) as mock_export:
                with patch("office_agent.office.finalizer.VisualReviewer") as mock_reviewer_cls:
                    mock_reviewer = mock_reviewer_cls.return_value
                    mock_reviewer.review.side_effect = [review_result, ReviewDeckResult(slides=[SlideReviewResult(slideIndex=0, issues=[])])]
                    with patch("office_agent.office.finalizer.apply_operations", return_value=(1, [])) as mock_apply:
                        result = finalize_presentation(
                            path,
                            spec=SPEC,
                            config=FinalizeConfig(enabled=True, model="gpt-4.1-mini", api_key="secret", max_rounds=2),
                        )
        self.assertEqual(result.summary.status, "completed")
        self.assertEqual(len(result.summary.rounds), 2)
        self.assertEqual(result.summary.issues_found, 1)
        self.assertEqual(result.summary.operations_applied, 1)
        self.assertEqual(mock_export.call_count, 2)
        mock_apply.assert_called_once()
