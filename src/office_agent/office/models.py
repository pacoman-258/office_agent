from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

from office_agent.schema import PresentationSpec


FinalizeStatus = Literal["skipped", "completed", "partial", "failed"]
FinalizeProvider = Literal["openai"]
IssueSeverity = Literal["low", "medium", "high"]
EditOperationType = Literal[
    "set_text",
    "set_font_size",
    "set_position",
    "set_size",
    "set_alignment",
    "set_paragraph_spacing",
    "hide_shape",
    "delete_shape",
    "duplicate_shape",
]
TextAlignment = Literal["left", "center", "right", "justify"]


class FinalizeConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    enabled: bool = False
    provider: FinalizeProvider = "openai"
    model: str | None = None
    api_key: str | None = None
    base_url: str = "https://api.openai.com/v1"
    max_rounds: int = Field(default=2, ge=1, le=5)


class ShapeSnapshot(BaseModel):
    model_config = ConfigDict(extra="forbid")

    shape_id: int = Field(alias="shapeId")
    name: str
    shape_type: str = Field(alias="shapeType")
    left: float
    top: float
    width: float
    height: float
    text: str = ""
    visible: bool = True


class SlideReviewInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    slide_index: int = Field(alias="slideIndex", ge=0)
    slide_title: str = Field(alias="slideTitle")
    slide_type: str = Field(alias="slideType")
    text_summary: list[str] = Field(alias="textSummary", default_factory=list)
    image_data_url: str = Field(alias="imageDataUrl")
    shapes: list[ShapeSnapshot] = Field(default_factory=list)


class SlideEditOperation(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: EditOperationType
    slide_index: int = Field(alias="slideIndex", ge=0)
    shape_id: int | None = Field(default=None, alias="shapeId")
    text: str | None = None
    font_size: float | None = Field(default=None, alias="fontSize", gt=1)
    left: float | None = None
    top: float | None = None
    width: float | None = Field(default=None, gt=1)
    height: float | None = Field(default=None, gt=1)
    alignment: TextAlignment | None = None
    space_before: float | None = Field(default=None, alias="spaceBefore", ge=0)
    space_after: float | None = Field(default=None, alias="spaceAfter", ge=0)
    duplicate_offset_x: float = Field(default=18.0, alias="duplicateOffsetX")
    duplicate_offset_y: float = Field(default=18.0, alias="duplicateOffsetY")


class SlideIssue(BaseModel):
    model_config = ConfigDict(extra="forbid")

    severity: IssueSeverity
    reason: str = Field(min_length=1)
    target_shape_id: int | None = Field(default=None, alias="targetShapeId")
    operations: list[SlideEditOperation] = Field(default_factory=list)


class SlideReviewResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    slide_index: int = Field(alias="slideIndex", ge=0)
    issues: list[SlideIssue] = Field(default_factory=list)


class ReviewDeckResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    slides: list[SlideReviewResult] = Field(default_factory=list)


class FinalizeRoundResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    round_index: int = Field(alias="roundIndex", ge=1)
    slides_reviewed: int = Field(alias="slidesReviewed", ge=0)
    issues_found: int = Field(alias="issuesFound", ge=0)
    operations_applied: int = Field(alias="operationsApplied", ge=0)
    warnings: list[str] = Field(default_factory=list)


class FinalizeSummary(BaseModel):
    model_config = ConfigDict(extra="forbid")

    enabled: bool = False
    status: FinalizeStatus = "skipped"
    provider: FinalizeProvider | None = None
    model: str | None = None
    rounds: list[FinalizeRoundResult] = Field(default_factory=list)
    issues_found: int = Field(default=0, alias="issuesFound", ge=0)
    operations_applied: int = Field(default=0, alias="operationsApplied", ge=0)
    warnings: list[str] = Field(default_factory=list)


class FinalizeResult(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True, extra="forbid")

    path: str
    summary: FinalizeSummary


class ReviewContext(BaseModel):
    model_config = ConfigDict(extra="forbid")

    spec: PresentationSpec
    slides: list[SlideReviewInput]
