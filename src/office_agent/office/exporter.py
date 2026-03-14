from __future__ import annotations

import base64
from contextlib import contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory

from office_agent.errors import OfficeAutomationError
from office_agent.office.models import ShapeSnapshot, SlideReviewInput
from office_agent.schema import PresentationSpec


THUMBNAIL_WIDTH = 1600
THUMBNAIL_HEIGHT = 900


@contextmanager
def powerpoint_session():
    try:
        import pythoncom
        from win32com import client
    except ImportError as exc:  # pragma: no cover - environment dependent
        raise OfficeAutomationError("PowerPoint automation requires pywin32 on Windows.") from exc

    pythoncom.CoInitialize()
    app = None
    try:
        app = client.DispatchEx("PowerPoint.Application")
        yield app
    except Exception as exc:  # pragma: no cover - environment dependent
        raise OfficeAutomationError("Microsoft PowerPoint COM automation is unavailable.") from exc
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


@contextmanager
def open_presentation(path: Path):
    with powerpoint_session() as app:
        presentation = None
        try:
            presentation = app.Presentations.Open(str(path), False, False, False)
            yield presentation
        except Exception as exc:  # pragma: no cover - environment dependent
            raise OfficeAutomationError(f"Failed to open presentation in PowerPoint: {exc}") from exc
        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception:
                    pass


def export_review_inputs(path: Path, spec: PresentationSpec) -> list[SlideReviewInput]:
    with TemporaryDirectory() as temp_dir:
        export_dir = Path(temp_dir)
        with open_presentation(path) as presentation:
            inputs: list[SlideReviewInput] = []
            for slide_index, slide_spec in enumerate(spec.slides, start=1):
                slide = presentation.Slides(slide_index)
                image_path = export_dir / f"slide-{slide_index - 1}.png"
                slide.Export(str(image_path), "PNG", THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT)
                inputs.append(
                    SlideReviewInput(
                        slideIndex=slide_index - 1,
                        slideTitle=slide_spec.title,
                        slideType=slide_spec.type,
                        textSummary=_slide_summary(slide_spec),
                        imageDataUrl=_png_to_data_url(image_path),
                        shapes=_collect_shapes(slide),
                    )
                )
            return inputs


def _collect_shapes(slide) -> list[ShapeSnapshot]:
    snapshots: list[ShapeSnapshot] = []
    for shape in slide.Shapes:
        text = ""
        try:
            if getattr(shape, "HasTextFrame", 0) and shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text or ""
        except Exception:
            text = ""
        try:
            visible = bool(shape.Visible)
        except Exception:
            visible = True
        snapshots.append(
            ShapeSnapshot(
                shapeId=int(shape.Id),
                name=str(getattr(shape, "Name", f"Shape-{shape.Id}")),
                shapeType=str(getattr(shape, "Type", "")),
                left=float(getattr(shape, "Left", 0.0)),
                top=float(getattr(shape, "Top", 0.0)),
                width=float(getattr(shape, "Width", 0.0)),
                height=float(getattr(shape, "Height", 0.0)),
                text=text.strip(),
                visible=visible,
            )
        )
    return snapshots


def _png_to_data_url(path: Path) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def _slide_summary(slide_spec) -> list[str]:
    summary: list[str] = []
    title = getattr(slide_spec, "title", None)
    if title:
        summary.append(title)
    for field_name in (
        "subtitle",
        "bullets",
        "key_points",
        "next_steps",
        "left_bullets",
        "right_bullets",
    ):
        value = getattr(slide_spec, field_name, None)
        if isinstance(value, list):
            summary.extend(str(item) for item in value)
        elif isinstance(value, str) and value.strip():
            summary.append(value.strip())
    if getattr(slide_spec, "type", None) == "comparison":
        summary.extend(getattr(slide_spec.left, "bullets", []))
        summary.extend(getattr(slide_spec.right, "bullets", []))
    if getattr(slide_spec, "type", None) == "timeline":
        summary.extend(f"{event.label}: {event.title}" for event in slide_spec.events)
    if getattr(slide_spec, "type", None) == "quote":
        summary.append(slide_spec.quote)
    if getattr(slide_spec, "type", None) == "table":
        summary.append(" | ".join(slide_spec.headers))
    return summary[:24]
