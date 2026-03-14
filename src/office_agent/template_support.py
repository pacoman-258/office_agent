from __future__ import annotations

from copy import deepcopy


PLACEHOLDER_PREFIX = "oa:"
SUPPORTED_PLACEHOLDER_ROLES = {"title", "subtitle", "body", "image", "caption"}


def duplicate_slide(prs, source_slide):
    layout = source_slide.slide_layout
    duplicated = prs.slides.add_slide(layout)

    for shape in list(duplicated.shapes):
        duplicated.shapes._spTree.remove(shape.element)

    for shape in source_slide.shapes:
        duplicated.shapes._spTree.insert_element_before(deepcopy(shape.element), "p:extLst")

    for rel in source_slide.part.rels.values():
        if "notesSlide" in rel.reltype:
            continue
        if rel.is_external:
            duplicated.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            duplicated.part.rels.get_or_add(rel.reltype, rel.target_part)

    return duplicated


def extract_placeholder_roles(slide) -> list[str]:
    roles = {role for shape in slide.shapes if (role := extract_shape_role(shape))}
    return sorted(roles)


def extract_shape_role(shape) -> str | None:
    for candidate in _shape_markers(shape):
        normalized = candidate.strip().lower()
        if not normalized.startswith(PLACEHOLDER_PREFIX):
            continue
        role = normalized.removeprefix(PLACEHOLDER_PREFIX)
        if role in SUPPORTED_PLACEHOLDER_ROLES:
            return role
    return None


def extract_slide_title_text(slide) -> str | None:
    title_shape = getattr(slide.shapes, "title", None)
    if title_shape is not None and getattr(title_shape, "has_text_frame", False):
        text = title_shape.text.strip()
        if text:
            return text

    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text = shape.text.strip()
            if text:
                return text
    return None


def _shape_markers(shape) -> list[str]:
    markers: list[str] = []
    name = getattr(shape, "name", None)
    if isinstance(name, str) and name.strip():
        markers.append(name)

    c_nv_pr = getattr(getattr(shape, "_element", None), "_nvXxPr", None)
    c_nv_pr = getattr(c_nv_pr, "cNvPr", None)
    if c_nv_pr is not None:
        for attribute in ("descr", "title"):
            value = c_nv_pr.get(attribute)
            if value:
                markers.append(value)

    return markers
