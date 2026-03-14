from __future__ import annotations

import base64
import tempfile
from pathlib import Path

from pptx import Presentation

from office_agent.api.models import TemplatePreviewResponse, TemplatePreviewSlide
from office_agent.errors import OfficeAgentError
from office_agent.template_support import TEMPLATE_CLEANUP_MODE, analyze_template_slide, extract_slide_title_text


THUMBNAIL_WIDTH = 1280
THUMBNAIL_HEIGHT = 720


class TemplatePreviewError(OfficeAgentError):
    """Raised when PPTX template preview generation fails."""


def build_template_preview(template_bytes: bytes, filename: str) -> TemplatePreviewResponse:
    if Path(filename or "template.pptx").suffix.lower() != ".pptx":
        raise TemplatePreviewError("Template preview only supports .pptx files.")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir) / "template.pptx"
        temp_path.write_bytes(template_bytes)
        try:
            presentation = Presentation(str(temp_path))
        except Exception as exc:
            raise TemplatePreviewError("The uploaded template is not a valid .pptx file.") from exc
        thumbnails = export_template_thumbnails(temp_path, Path(temp_dir))
        slides: list[TemplatePreviewSlide] = []
        for index, slide in enumerate(presentation.slides):
            thumbnail_path = thumbnails.get(index)
            if thumbnail_path is None or not thumbnail_path.exists():
                raise TemplatePreviewError("Failed to generate thumbnails for all template slides.")
            analysis = analyze_template_slide(slide)
            slides.append(
                TemplatePreviewSlide(
                    index=index,
                    thumbnailDataUrl=png_to_data_url(thumbnail_path),
                    titleText=extract_slide_title_text(slide),
                    placeholderRoles=analysis.placeholder_roles,
                )
            )
        return TemplatePreviewResponse(slides=slides, cleanupMode=TEMPLATE_CLEANUP_MODE)


def export_template_thumbnails(template_path: Path, output_dir: Path) -> dict[int, Path]:
    try:
        import pythoncom
        from win32com import client
    except ImportError as exc:  # pragma: no cover - environment dependent
        raise TemplatePreviewError(
            "Template preview requires pywin32 and Microsoft PowerPoint on Windows."
        ) from exc

    export_dir = output_dir / "thumbnails"
    export_dir.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = client.DispatchEx("PowerPoint.Application")
        app.Visible = 0
        presentation = app.Presentations.Open(str(template_path), False, False, False)
        results: dict[int, Path] = {}
        for index, slide in enumerate(presentation.Slides):
            file_path = export_dir / f"slide-{index}.png"
            slide.Export(str(file_path), "PNG", THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT)
            results[index] = file_path
        return results
    except Exception as exc:  # pragma: no cover - environment dependent
        raise TemplatePreviewError(
            "PowerPoint could not generate template thumbnails. Confirm Microsoft PowerPoint is installed."
        ) from exc
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def png_to_data_url(path: Path) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"
