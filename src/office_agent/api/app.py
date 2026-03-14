from __future__ import annotations

import json
from json import JSONDecodeError
from urllib.parse import quote

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import JSONResponse, Response
from pydantic import ValidationError

from office_agent.api.models import (
    GenerateSpecRequest,
    HealthResponse,
    RenderPresentationRequest,
    TemplatePreviewResponse,
)
from office_agent.errors import OfficeAgentError
from office_agent.schema import PresentationSpec
from office_agent.services import (
    generate_spec_from_prompt,
    preview_template_artifact,
    render_presentation_artifact,
)


def create_app() -> FastAPI:
    app = FastAPI(title="office-agent API", version="0.1.0")

    @app.exception_handler(OfficeAgentError)
    def handle_application_error(_, exc: OfficeAgentError) -> JSONResponse:
        return JSONResponse(status_code=400, content={"detail": str(exc)})

    @app.exception_handler(ValidationError)
    def handle_validation_error(_, exc: ValidationError) -> JSONResponse:
        return JSONResponse(status_code=422, content={"detail": exc.errors()})

    @app.get("/api/health", response_model=HealthResponse)
    def health() -> HealthResponse:
        return HealthResponse(providers=["openai", "ollama"], themes=["default", "executive", "editorial"])

    @app.post("/api/specs", response_model=PresentationSpec)
    def generate_spec(request: GenerateSpecRequest) -> PresentationSpec:
        config = request.to_app_config()
        config.validate()
        return generate_spec_from_prompt(request.prompt, config, template_mapping=request.template_mapping)

    @app.post("/api/templates/preview", response_model=TemplatePreviewResponse)
    async def preview_template(template: UploadFile = File(...)) -> TemplatePreviewResponse:
        template_bytes = await template.read()
        try:
            return preview_template_artifact(template_bytes, template.filename or "template.pptx")
        finally:
            await template.close()

    @app.post("/api/presentations")
    async def render_presentation(request: Request) -> Response:
        render_request, template_upload = await parse_render_request(request)
        try:
            template_bytes = await template_upload.read() if template_upload is not None else None
            artifact = render_presentation_artifact(
                render_request.spec,
                filename=render_request.filename,
                template_bytes=template_bytes,
                template_filename=template_upload.filename if template_upload is not None else None,
                finalize_config=render_request.finalize.to_finalize_config() if render_request.finalize is not None else None,
            )
        finally:
            if template_upload is not None:
                await template_upload.close()
        quoted_filename = quote(artifact.filename)
        headers = {
            "Content-Disposition": f'attachment; filename="presentation.pptx"; filename*=UTF-8\'\'{quoted_filename}',
        }
        if artifact.warnings:
            headers["X-Office-Agent-Warnings"] = quote(json.dumps(artifact.warnings, ensure_ascii=False))
        if artifact.finalize_summary is not None:
            headers["X-Office-Agent-Finalize"] = quote(
                str(artifact.finalize_summary.model_dump_json(by_alias=True, exclude_none=True))
            )
        return Response(
            content=artifact.content,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers=headers,
        )

    return app


async def parse_render_request(request: Request) -> tuple[RenderPresentationRequest, UploadFile | None]:
    content_type = request.headers.get("content-type", "")
    if content_type.startswith("application/json"):
        payload = await request.json()
        return RenderPresentationRequest.model_validate(payload), None

    if content_type.startswith("multipart/form-data"):
        form = await request.form()
        raw_payload = form.get("payload")
        if not isinstance(raw_payload, str):
            raise OfficeAgentError("Missing payload field in multipart render request.")
        try:
            payload = json.loads(raw_payload)
        except JSONDecodeError as exc:
            raise OfficeAgentError("Invalid payload JSON in multipart render request.") from exc
        template = form.get("template")
        upload = template if isinstance(template, UploadFile) or hasattr(template, "read") else None
        return RenderPresentationRequest.model_validate(payload), upload

    raise OfficeAgentError("Unsupported content type for presentation rendering.")
