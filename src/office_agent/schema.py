from __future__ import annotations

import re
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, HttpUrl, TypeAdapter, field_validator, model_validator


HEX_COLOR_PATTERN = re.compile(r"^#[0-9A-Fa-f]{6}$")
TemplatePart = Literal["opening", "agenda", "content", "closing"]


class CustomThemeSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    primary_color: str | None = None
    accent_color: str | None = None
    background_color: str | None = None
    heading_font: str | None = None
    body_font: str | None = None
    cover_style: Literal["band", "centered", "minimal"] | None = None

    @field_validator("primary_color", "accent_color", "background_color")
    @classmethod
    def validate_hex_color(cls, value: str | None) -> str | None:
        if value is None:
            return None
        if not HEX_COLOR_PATTERN.match(value):
            raise ValueError("theme colors must use #RRGGBB format")
        return value

    @field_validator("heading_font", "body_font")
    @classmethod
    def validate_font_name(cls, value: str | None) -> str | None:
        if value is None:
            return None
        trimmed = value.strip()
        if not trimmed:
            raise ValueError("font names must be non-empty when provided")
        return trimmed


class ThemeSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    preset: Literal["default", "executive", "editorial"] = "default"
    custom: CustomThemeSpec | None = None


class TemplateSelectionSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    opening: int = Field(ge=0)
    agenda: int = Field(ge=0)
    content: int = Field(ge=0)
    closing: int = Field(ge=0)


class BaseSlideSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    part: TemplatePart
    title: str


class TitleSlideSpec(BaseSlideSpec):
    type: Literal["title"]
    subtitle: str | None = None


class SectionSlideSpec(BaseSlideSpec):
    type: Literal["section"]
    subtitle: str | None = None


class BulletsSlideSpec(BaseSlideSpec):
    type: Literal["bullets"]
    bullets: list[str] = Field(min_length=1)


class TwoColumnSlideSpec(BaseSlideSpec):
    type: Literal["two_column"]
    left_title: str | None = None
    left_bullets: list[str] = Field(min_length=1)
    right_title: str | None = None
    right_bullets: list[str] = Field(min_length=1)


class ImageSlideSpec(BaseSlideSpec):
    type: Literal["image"]
    image: str
    caption: str | None = None
    bullets: list[str] = Field(default_factory=list)

    @field_validator("image")
    @classmethod
    def validate_image_source(cls, value: str) -> str:
        if value.startswith(("http://", "https://")):
            TypeAdapter(HttpUrl).validate_python(value)
        elif not value.strip():
            raise ValueError("image must be a non-empty local path or HTTP(S) URL")
        return value


class TimelineItemSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    label: str
    title: str
    detail: str | None = None


class TimelineSlideSpec(BaseSlideSpec):
    type: Literal["timeline"]
    events: list[TimelineItemSpec] = Field(min_length=2)


class QuoteSlideSpec(BaseSlideSpec):
    type: Literal["quote"]
    quote: str
    attribution: str | None = None
    source: str | None = None


class ComparisonColumnSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str
    bullets: list[str] = Field(min_length=1)


class ComparisonSlideSpec(BaseSlideSpec):
    type: Literal["comparison"]
    left: ComparisonColumnSpec
    right: ComparisonColumnSpec


class SummarySlideSpec(BaseSlideSpec):
    type: Literal["summary"]
    key_points: list[str] = Field(min_length=1)
    next_steps: list[str] = Field(default_factory=list)


class TableSlideSpec(BaseSlideSpec):
    type: Literal["table"]
    headers: list[str] = Field(min_length=1)
    rows: list[list[str]] = Field(min_length=1)

    @model_validator(mode="after")
    def validate_row_lengths(self) -> "TableSlideSpec":
        expected_columns = len(self.headers)
        for row in self.rows:
            if len(row) != expected_columns:
                raise ValueError("each table row must match the number of headers")
        return self


SlideSpec = Annotated[
    TitleSlideSpec
    | SectionSlideSpec
    | BulletsSlideSpec
    | TwoColumnSlideSpec
    | ImageSlideSpec
    | TimelineSlideSpec
    | QuoteSlideSpec
    | ComparisonSlideSpec
    | SummarySlideSpec
    | TableSlideSpec,
    Field(discriminator="type"),
]


class PresentationSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str
    theme: ThemeSpec = Field(default_factory=ThemeSpec)
    template: TemplateSelectionSpec | None = None
    slides: list[SlideSpec] = Field(min_length=1)
