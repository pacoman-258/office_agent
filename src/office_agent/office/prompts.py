from __future__ import annotations

from office_agent.office.models import ReviewContext


SYSTEM_PROMPT = """
You are reviewing rendered PowerPoint slides and proposing safe visual-fix operations.
Return only valid JSON.

Goal:
- Find layout defects that are clearly visible on the slide screenshot.
- Focus on text overflow, overlap, clipping, edge collisions, spacing problems, unbalanced columns, and leftover non-branding artifacts.
- Do not invent content changes unless needed to repair visible layout issues.
- Prefer the smallest safe fix.

Rules:
- Output this exact top-level shape:
  {"slides": [{"slideIndex": 0, "issues": [{"severity": "low|medium|high", "reason": "...", "targetShapeId": 1, "operations": [...]}]}]}
- Use only these operations:
  set_text, set_font_size, set_position, set_size, set_alignment, set_paragraph_spacing, hide_shape, delete_shape, duplicate_shape
- Every operation must include slideIndex.
- Shape-targeting operations must use an existing shapeId from the provided shape inventory.
- Never output VBA, arbitrary code, or unsupported operations.
- Do not delete branding, footer, page-number, or logo-like elements unless the issue explicitly says they are covering generated content.
- If a slide looks acceptable, return that slide with an empty issues array.
- Prefer one or two operations per issue.
""".strip()


def build_review_messages(context: ReviewContext) -> list[dict[str, object]]:
    user_content: list[dict[str, object]] = [
        {
            "type": "text",
            "text": (
                "Review this rendered presentation and propose structured fix operations.\n\n"
                "Presentation spec:\n"
                f"{context.spec.model_dump_json(indent=2)}\n\n"
                "For each slide below, use the screenshot plus shape inventory."
            ),
        }
    ]
    for slide in context.slides:
        user_content.append(
            {
                "type": "text",
                "text": (
                    f"Slide {slide.slide_index} ({slide.slide_type})\n"
                    f"Title: {slide.slide_title}\n"
                    f"Text summary: {slide.text_summary}\n"
                    f"Shape inventory: {slide.model_dump_json(indent=2, by_alias=True, exclude={'image_data_url'})}"
                ),
            }
        )
        user_content.append({"type": "image_url", "image_url": {"url": slide.image_data_url}})

    return [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]
