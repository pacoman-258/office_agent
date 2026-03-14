from __future__ import annotations

import os
from pathlib import Path

import requests

from office_agent.errors import OfficeAutomationError
from office_agent.office.editor import apply_operations
from office_agent.office.exporter import export_review_inputs
from office_agent.office.models import (
    FinalizeConfig,
    FinalizeResult,
    FinalizeRoundResult,
    FinalizeSummary,
    ReviewContext,
    SlideEditOperation,
)
from office_agent.office.reviewer import VisualReviewer
from office_agent.schema import PresentationSpec


def finalize_presentation(
    path: str | Path,
    *,
    spec: PresentationSpec | None = None,
    config: FinalizeConfig | None = None,
    session: requests.Session | None = None,
) -> FinalizeResult:
    output_path = Path(path)
    finalize_config = config or finalize_config_from_env()
    summary = FinalizeSummary(
        enabled=finalize_config.enabled,
        provider=finalize_config.provider if finalize_config.enabled else None,
        model=finalize_config.model if finalize_config.enabled else None,
    )

    if not finalize_config.enabled:
        summary.status = "skipped"
        return FinalizeResult(path=str(output_path), summary=summary)

    if spec is None:
        summary.status = "skipped"
        summary.warnings.append("Visual finalization skipped because the presentation spec was not provided.")
        return FinalizeResult(path=str(output_path), summary=summary)

    if finalize_config.provider != "openai":
        summary.status = "skipped"
        summary.warnings.append(f"Visual finalization only supports OpenAI-compatible providers in this version, got '{finalize_config.provider}'.")
        return FinalizeResult(path=str(output_path), summary=summary)

    if not finalize_config.api_key:
        summary.status = "skipped"
        summary.warnings.append("Visual finalization skipped because no OpenAI-compatible API key was provided.")
        return FinalizeResult(path=str(output_path), summary=summary)

    if not finalize_config.model:
        summary.status = "skipped"
        summary.warnings.append("Visual finalization skipped because no visual review model was configured.")
        return FinalizeResult(path=str(output_path), summary=summary)

    reviewer = VisualReviewer(finalize_config, session=session)
    pending_issue_count = 0
    try:
        for round_index in range(1, finalize_config.max_rounds + 1):
            slides = export_review_inputs(output_path, spec)
            review_result = reviewer.review(ReviewContext(spec=spec, slides=slides))
            issues_found = sum(len(slide.issues) for slide in review_result.slides)
            operations = [
                operation
                for slide in review_result.slides
                for issue in slide.issues
                for operation in issue.operations
            ]
            if operations:
                applied_operations, round_warnings = apply_operations(output_path, operations)
            else:
                applied_operations, round_warnings = 0, []
            summary.rounds.append(
                FinalizeRoundResult(
                    roundIndex=round_index,
                    slidesReviewed=len(slides),
                    issuesFound=issues_found,
                    operationsApplied=applied_operations,
                    warnings=round_warnings,
                )
            )
            summary.issues_found += issues_found
            summary.operations_applied += applied_operations
            summary.warnings.extend(round_warnings)
            pending_issue_count = issues_found
            if issues_found == 0 or applied_operations == 0:
                break
        summary.status = "completed" if pending_issue_count == 0 else "partial"
        return FinalizeResult(path=str(output_path), summary=summary)
    except OfficeAutomationError as exc:
        summary.status = "failed"
        summary.warnings.append(str(exc))
        return FinalizeResult(path=str(output_path), summary=summary)


def finalize_config_from_env() -> FinalizeConfig:
    model = os.getenv("OFFICE_AGENT_FINALIZE_MODEL")
    return FinalizeConfig(
        enabled=_env_bool("OFFICE_AGENT_FINALIZE_ENABLED", False),
        provider="openai",
        model=model.strip() if model else None,
        api_key=(os.getenv("OPENAI_VISION_API_KEY") or os.getenv("OPENAI_API_KEY") or "").strip() or None,
        base_url=(os.getenv("OPENAI_VISION_BASE_URL") or os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").rstrip("/"),
        max_rounds=_env_int("OFFICE_AGENT_FINALIZE_MAX_ROUNDS", 2),
    )


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return max(1, min(5, int(raw)))
    except ValueError:
        return default
