from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from office_agent.errors import OfficeAutomationError
from office_agent.office.exporter import open_presentation
from office_agent.office.models import SlideEditOperation


ALIGNMENT_MAP = {
    "left": 1,
    "center": 2,
    "right": 3,
    "justify": 4,
}


def apply_operations(path: Path, operations: Iterable[SlideEditOperation]) -> tuple[int, list[str]]:
    operation_list = list(operations)
    if not operation_list:
        return 0, []

    warnings: list[str] = []
    applied = 0
    with open_presentation(path) as presentation:
        for operation in operation_list:
            try:
                if _apply_operation(presentation, operation):
                    applied += 1
            except OfficeAutomationError as exc:
                warnings.append(str(exc))
        try:
            presentation.Save()
        except Exception as exc:  # pragma: no cover - environment dependent
            raise OfficeAutomationError(f"Failed to save PowerPoint changes: {exc}") from exc
    return applied, warnings


def _apply_operation(presentation, operation: SlideEditOperation) -> bool:
    slide = presentation.Slides(operation.slide_index + 1)
    shape = _find_shape(slide, operation.shape_id) if operation.shape_id is not None else None

    match operation.type:
        case "set_text":
            if shape is None:
                raise OfficeAutomationError("set_text requires a valid shapeId.")
            _ensure_text_shape(shape, operation.type)
            shape.TextFrame.TextRange.Text = operation.text or ""
            return True
        case "set_font_size":
            if shape is None:
                raise OfficeAutomationError("set_font_size requires a valid shapeId.")
            _ensure_text_shape(shape, operation.type)
            shape.TextFrame.TextRange.Font.Size = operation.font_size
            return True
        case "set_position":
            if shape is None:
                raise OfficeAutomationError("set_position requires a valid shapeId.")
            if operation.left is not None:
                shape.Left = operation.left
            if operation.top is not None:
                shape.Top = operation.top
            return True
        case "set_size":
            if shape is None:
                raise OfficeAutomationError("set_size requires a valid shapeId.")
            if operation.width is not None:
                shape.Width = operation.width
            if operation.height is not None:
                shape.Height = operation.height
            return True
        case "set_alignment":
            if shape is None:
                raise OfficeAutomationError("set_alignment requires a valid shapeId.")
            _ensure_text_shape(shape, operation.type)
            alignment = ALIGNMENT_MAP.get(operation.alignment or "")
            if alignment is None:
                raise OfficeAutomationError(f"Unsupported text alignment: {operation.alignment}")
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = alignment
            return True
        case "set_paragraph_spacing":
            if shape is None:
                raise OfficeAutomationError("set_paragraph_spacing requires a valid shapeId.")
            _ensure_text_shape(shape, operation.type)
            paragraph_format = shape.TextFrame.TextRange.ParagraphFormat
            if operation.space_before is not None:
                paragraph_format.SpaceBefore = operation.space_before
            if operation.space_after is not None:
                paragraph_format.SpaceAfter = operation.space_after
            return True
        case "hide_shape":
            if shape is None:
                raise OfficeAutomationError("hide_shape requires a valid shapeId.")
            shape.Visible = False
            return True
        case "delete_shape":
            if shape is None:
                raise OfficeAutomationError("delete_shape requires a valid shapeId.")
            shape.Delete()
            return True
        case "duplicate_shape":
            if shape is None:
                raise OfficeAutomationError("duplicate_shape requires a valid shapeId.")
            duplicate = shape.Duplicate()
            duplicated_shape = duplicate.Item(1) if hasattr(duplicate, "Item") else duplicate
            duplicated_shape.Left = shape.Left + operation.duplicate_offset_x
            duplicated_shape.Top = shape.Top + operation.duplicate_offset_y
            return True
    raise OfficeAutomationError(f"Unsupported edit operation: {operation.type}")


def _find_shape(slide, shape_id: int | None):
    if shape_id is None:
        return None
    for shape in slide.Shapes:
        try:
            if int(shape.Id) == shape_id:
                return shape
        except Exception:
            continue
    raise OfficeAutomationError(f"Could not find shapeId {shape_id} on slide.")


def _ensure_text_shape(shape, operation_type: str) -> None:
    try:
        if not getattr(shape, "HasTextFrame", 0):
            raise OfficeAutomationError(f"{operation_type} requires a text shape.")
    except Exception as exc:
        if isinstance(exc, OfficeAutomationError):
            raise
        raise OfficeAutomationError(f"{operation_type} requires a text shape.") from exc
