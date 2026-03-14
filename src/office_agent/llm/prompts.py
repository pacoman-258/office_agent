from __future__ import annotations

from office_agent.schema import TemplateSelectionSpec, ThemeSpec


BASE_SLIDE_RULES = """
You are generating a PowerPoint specification as JSON.
Return only valid JSON.

Use this top-level structure:
{
  "title": "presentation title",
  "theme": {
    "preset": "default | executive | editorial",
    "custom": {
      "primary_color": "#RRGGBB",
      "accent_color": "#RRGGBB",
      "background_color": "#RRGGBB",
      "heading_font": "Font Name",
      "body_font": "Font Name",
      "cover_style": "band | centered | minimal"
    }
  },
  "template": {
    "opening": 0,
    "agenda": 1,
    "content": 2,
    "closing": 3
  },
  "slides": [
    {"type": "title", "part": "opening", "title": "...", "subtitle": "..."},
    {"type": "section", "part": "agenda", "title": "...", "subtitle": "..."},
    {"type": "bullets", "part": "content", "title": "...", "bullets": ["...", "..."]},
    {
      "type": "two_column",
      "part": "content",
      "title": "...",
      "left_title": "...",
      "left_bullets": ["...", "..."],
      "right_title": "...",
      "right_bullets": ["...", "..."]
    },
    {
      "type": "image",
      "part": "content",
      "title": "...",
      "image": "https://example.com/image.png or local/path.png",
      "caption": "...",
      "bullets": ["...", "..."]
    },
    {
      "type": "timeline",
      "part": "content",
      "title": "...",
      "events": [{"label": "Q1", "title": "...", "detail": "..."}]
    },
    {
      "type": "quote",
      "part": "content",
      "title": "...",
      "quote": "...",
      "attribution": "...",
      "source": "..."
    },
    {
      "type": "comparison",
      "part": "content",
      "title": "...",
      "left": {"title": "...", "bullets": ["..."]},
      "right": {"title": "...", "bullets": ["..."]}
    },
    {
      "type": "summary",
      "part": "closing",
      "title": "...",
      "key_points": ["..."],
      "next_steps": ["..."]
    },
    {
      "type": "table",
      "part": "content",
      "title": "...",
      "headers": ["...", "..."],
      "rows": [["...", "..."], ["...", "..."]]
    }
  ]
}

Rules:
- Use only the ten supported slide types.
- Every slide must include a `part` value: opening, agenda, content, or closing.
- Do not add extra keys.
- Match the requested theme object exactly.
- If a template mapping is provided, match the requested template object exactly.
- Produce 3 to 8 slides unless the user explicitly asks for another size.
- Prefer concise slide text suitable for a presentation.
- Keep output language aligned with the user's prompt.
- Use table slides only for compact structured data. Do not invent charts.
- Assign parts consistently:
  - The first slide should usually be `opening`.
  - A table-of-contents or agenda slide should be `agenda`.
  - A closing summary, conclusion, thank-you, or next-steps slide should be `closing`.
  - Other slides should be `content`.
""".strip()


def build_messages(
    user_prompt: str,
    requested_theme: ThemeSpec,
    template_mapping: TemplateSelectionSpec | None = None,
    previous_error: str | None = None,
) -> list[dict[str, str]]:
    instructions = [
        "Use this exact theme object in your JSON response:",
        requested_theme.model_dump_json(indent=2),
    ]
    if template_mapping is not None:
        instructions.extend(
            [
                "",
                "Use this exact template object in your JSON response:",
                template_mapping.model_dump_json(indent=2),
            ]
        )
    else:
        instructions.extend(
            [
                "",
                "If no template mapping is provided, set `template` to null or omit it.",
            ]
        )

    messages = [
        {"role": "system", "content": BASE_SLIDE_RULES},
        {
            "role": "user",
            "content": "\n".join(instructions) + f"\n\nUser request:\n{user_prompt}",
        },
    ]
    if previous_error:
        messages.append(
            {
                "role": "user",
                "content": (
                    "Your previous response could not be parsed or validated. "
                    f"Fix the JSON and return only JSON.\nError: {previous_error}"
                ),
            }
        )
    return messages
