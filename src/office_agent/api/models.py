from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

from office_agent.config import AppConfig
from office_agent.schema import PresentationSpec, TemplateSelectionSpec, ThemeSpec


Provider = Literal["openai", "ollama"]
ThemePreset = Literal["default", "executive", "editorial"]


class RuntimeProviderConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    provider: Provider
    model: str = Field(min_length=1)
    theme: ThemeSpec = Field(default_factory=ThemeSpec)
    api_key: str | None = Field(default=None, alias="apiKey")
    openai_base_url: str | None = Field(default=None, alias="openaiBaseUrl")
    ollama_base_url: str | None = Field(default=None, alias="ollamaBaseUrl")

    def to_app_config(self) -> AppConfig:
        return AppConfig.from_env().with_overrides(
            provider=self.provider,
            model=self.model,
            theme=self.theme,
            openai_api_key=self.api_key,
            openai_base_url=self.openai_base_url,
            ollama_base_url=self.ollama_base_url,
        )


class GenerateSpecRequest(RuntimeProviderConfig):
    prompt: str = Field(min_length=1)
    template_mapping: TemplateSelectionSpec | None = Field(default=None, alias="templateMapping")


class RenderPresentationRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    spec: PresentationSpec
    filename: str = Field(min_length=1)


class TemplatePreviewSlide(BaseModel):
    model_config = ConfigDict(extra="forbid")

    index: int = Field(ge=0)
    thumbnail_data_url: str = Field(alias="thumbnailDataUrl")
    title_text: str | None = Field(default=None, alias="titleText")
    placeholder_roles: list[str] = Field(default_factory=list, alias="placeholderRoles")


class TemplatePreviewResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")

    slides: list[TemplatePreviewSlide]


class HealthResponse(BaseModel):
    status: Literal["ok"] = "ok"
    providers: list[Provider]
    themes: list[ThemePreset]
