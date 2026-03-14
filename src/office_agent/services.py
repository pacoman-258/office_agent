from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

from office_agent.config import AppConfig
from office_agent.errors import RenderError
from office_agent.llm import generate_presentation_spec
from office_agent.office import finalize_presentation
from office_agent.office.models import FinalizeConfig, FinalizeSummary
from office_agent.renderer import PresentationRenderer, RenderResult
from office_agent.schema import PresentationSpec, TemplateSelectionSpec
from office_agent.template_preview import build_template_preview


@dataclass(frozen=True)
class PresentationArtifact:
    filename: str
    content: bytes
    warnings: list[str]
    finalize_summary: FinalizeSummary | None = None


INVALID_WINDOWS_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def generate_spec_from_prompt(
    prompt: str,
    config: AppConfig,
    template_mapping: TemplateSelectionSpec | None = None,
) -> PresentationSpec:
    return generate_presentation_spec(prompt, config.model, config, template_mapping=template_mapping)


def render_presentation(
    spec: PresentationSpec,
    out_path: str | Path,
    template_path: str | Path | None = None,
    finalize_config: FinalizeConfig | None = None,
) -> RenderResult:
    renderer = PresentationRenderer()
    result = renderer.render(spec, out_path, template_path=template_path)
    finalize_result = finalize_presentation(result.path, spec=spec, config=finalize_config)
    warnings = list(result.warnings)
    warnings.extend(finalize_result.summary.warnings)
    return RenderResult(path=Path(finalize_result.path), warnings=warnings, finalize_summary=finalize_result.summary)


def render_presentation_artifact(
    spec: PresentationSpec,
    filename: str,
    template_bytes: bytes | None = None,
    template_filename: str | None = None,
    finalize_config: FinalizeConfig | None = None,
) -> PresentationArtifact:
    download_name = normalize_download_filename(filename)
    temp_name = "presentation.pptx"
    with tempfile.TemporaryDirectory() as temp_dir:
        out_path = Path(temp_dir) / temp_name
        template_path: Path | None = None
        if template_bytes is not None:
            template_suffix = Path(template_filename or "template.pptx").suffix.lower()
            if template_suffix != ".pptx":
                raise RenderError("Template file must use the .pptx format.")
            template_path = Path(temp_dir) / "uploaded-template.pptx"
            template_path.write_bytes(template_bytes)
        result = render_presentation(spec, out_path, template_path=template_path, finalize_config=finalize_config)
        content = result.path.read_bytes()
    return PresentationArtifact(
        filename=download_name,
        content=content,
        warnings=result.warnings,
        finalize_summary=result.finalize_summary,
    )


def normalize_download_filename(filename: str) -> str:
    normalized_name = Path(filename).name.strip() or "presentation"
    if not normalized_name.lower().endswith(".pptx"):
        normalized_name = f"{normalized_name}.pptx"
    stem = Path(normalized_name).stem.strip().rstrip(". ")
    suffix = Path(normalized_name).suffix or ".pptx"
    cleaned_stem = INVALID_WINDOWS_FILENAME_CHARS.sub("_", stem).strip() or "presentation"
    return f"{cleaned_stem}{suffix}"


def preview_template_artifact(template_bytes: bytes, filename: str):
    return build_template_preview(template_bytes, filename)
