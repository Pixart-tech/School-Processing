"""Utilities for generating personalised ID card PDFs from SVG templates."""
from __future__ import annotations
import argparse
import csv
import math
import re
import shutil
from pathlib import Path
from typing import Dict, Iterable, Iterator, Optional, Sequence, Tuple, NamedTuple
from PIL import ImageFont
from xml.dom.minidom import Document, Element, Node, parse

from doc_maker import callInkscape


DEFAULT_TEMPLATE_ROOT = Path(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL ID CARD SRC")
DEFAULT_OUTPUT_ROOT = Path("ID Cards")
DEFAULT_PHOTO_ROOT = Path(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL_PHOTOS")


class TemplateNotFoundError(FileNotFoundError):
    """Raised when the required SVG templates for a school cannot be located."""


def _is_missing(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float):
        return math.isnan(value)
    return False


def _normalise_string(value: object, default: str = "") -> str:
    if _is_missing(value):
        return default
    value_str = str(value).strip()
    return value_str if value_str else default


def _ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _sanitize_filename_component(value: str, fallback: str) -> str:
    value = _normalise_string(value)
    if not value:
        value = fallback
    sanitized = re.sub(r"[^A-Za-z0-9]+", "_", value)
    sanitized = re.sub(r"_+", "_", sanitized).strip("_")
    return sanitized or fallback


def _build_child_output_base(first_name: str, last_name: str, school_name: str) -> str:
    parts = []
    if first_name:
        parts.append(_sanitize_filename_component(first_name, "student"))
    else:
        parts.append("student")
    if last_name:
        parts.append(_sanitize_filename_component(last_name, ""))
    if school_name:
        parts.append(_sanitize_filename_component(school_name, "school"))
    else:
        parts.append("school")
    base = "_".join(part for part in parts if part)
    return base or "student_school"


_title_case_regex = re.compile(r"\b\w+\b")


def custom_title_case(value: str) -> str:
    """Replicate the Illustrator script's custom title case logic."""

    def _transform(match: re.Match[str]) -> str:
        word = match.group(0)
        if len(word) <= 2:
            return word.upper()
        return word[0].upper() + word[1:].lower()

    lower_value = value.lower()
    return _title_case_regex.sub(_transform, lower_value)


_branch_cleanup_regex = re.compile(r"(kids|castle|feather|touch)", re.IGNORECASE)


def clean_branch_name(value: str) -> str:
    return _branch_cleanup_regex.sub("", value).strip()


FONT_SIZE_PATTERN = re.compile(r"font-size\s*:\s*([0-9.]+)px", re.IGNORECASE)
FONT_FAMILY_PATTERN = re.compile(r"font-family\s*:\s*([^;]+)", re.IGNORECASE)
_TRANSLATE_RE = re.compile(r"translate\s*\(\s*([^)]+)\)", re.IGNORECASE)
_COORDINATE_SPLIT_RE = re.compile(r"[\s,]+")

MIN_FONT_SIZE = 4.6
MULTILINE_MIN_FONT_SIZE = 4.6
BASELINE_SPACING_SCALE = 0.92


CENTER_ALIGNED_GROUPS = {"name", "fname", "mname", "fcontact", "mcontact"}
LEFT_ALIGNED_GROUPS = {"grade"}


_ALIGNMENT_CHAR_MAP = {"L": "left", "M": "center", "R": "right"}
_ALIGNMENT_WORD_MAP = {
    "LEFT": "left",
    "RIGHT": "right",
    "CENTER": "center",
    "CENTRE": "center",
    "MIDDLE": "center",
}
_ALIGNMENT_PREFIX_RE = re.compile(r"^\s*([LMR])(?:\b|[_\-\s:])", re.IGNORECASE)
_ALIGNMENT_SUFFIX_RE = re.compile(r"(?:^|[_\-\s:])([LMR])\s*$", re.IGNORECASE)


def _interpret_alignment_token(value: str) -> Optional[str]:
    if not value:
        return None

    cleaned = value.strip()
    if not cleaned:
        return None

    upper = cleaned.upper()
    if upper in _ALIGNMENT_WORD_MAP:
        return _ALIGNMENT_WORD_MAP[upper]

    if len(cleaned) == 1:
        return _ALIGNMENT_CHAR_MAP.get(upper)

    prefix_match = _ALIGNMENT_PREFIX_RE.match(cleaned)
    if prefix_match:
        return _ALIGNMENT_CHAR_MAP.get(prefix_match.group(1).upper())

    suffix_match = _ALIGNMENT_SUFFIX_RE.search(cleaned)
    if suffix_match:
        return _ALIGNMENT_CHAR_MAP.get(suffix_match.group(1).upper())

    tokens = re.split(r"[^A-Za-z]+", upper)
    for token in tokens:
        if not token:
            continue
        if token in _ALIGNMENT_WORD_MAP:
            return _ALIGNMENT_WORD_MAP[token]
        if len(token) == 1 and token in _ALIGNMENT_CHAR_MAP:
            return _ALIGNMENT_CHAR_MAP[token]

    return None


def _resolve_layer_alignment(element: Element) -> Optional[str]:
    current: Optional[Element] = element
    while isinstance(current, Element):
        tag_name = current.tagName.lower() if hasattr(current, "tagName") else ""
        if tag_name in {"g", "text"}:
            for attribute in ("inkscape:label", "id"):
                if current.hasAttribute(attribute):
                    alignment = _interpret_alignment_token(current.getAttribute(attribute))
                    if alignment:
                        return alignment

        parent = current.parentNode
        while parent is not None and parent.nodeType != Node.ELEMENT_NODE:
            parent = parent.parentNode

        current = parent if isinstance(parent, Element) else None

    return None


def _set_text(element: Element, text: str) -> None:
    while element.firstChild:
        element.removeChild(element.firstChild)
    element.appendChild(element.ownerDocument.createTextNode(text))


def _set_multiline_text(
    element: Element, lines: Sequence[str], *, line_height: Optional[float] = None
) -> None:
    document = element.ownerDocument
    base_x = element.getAttribute("x") if element.hasAttribute("x") else ""

    while element.firstChild:
        element.removeChild(element.firstChild)

    if not lines:
        lines = [""]

    element.appendChild(document.createTextNode(lines[0]))

    for line in lines[1:]:
        tspan = document.createElement("tspan")
        if base_x:
            tspan.setAttribute("x", base_x)
        if line_height is not None:
            tspan.setAttribute("dy", f"{_format_float(line_height)}px")
        else:
            tspan.setAttribute("dy", "1em")
        tspan.appendChild(document.createTextNode(line))
        element.appendChild(tspan)


def _format_float(value: float) -> str:
    return ("{:.4f}".format(value)).rstrip("0").rstrip(".")


def _extract_template_lines(element: Element) -> Sequence[str]:
    lines = []
    current = []

    for node in element.childNodes:
        if node.nodeType == Node.TEXT_NODE:
            current.append(node.data or "")
        elif node.nodeType == Node.ELEMENT_NODE:
            if current:
                combined = "".join(current).strip()
                if combined:
                    lines.append(combined)
                current = []
            lines.extend(_extract_template_lines(node))

    if current:
        combined = "".join(current).strip()
        if combined:
            lines.append(combined)

    return lines


def _set_font_size(element: Element, font_size: float) -> None:
    style = element.getAttribute("style") or ""
    if FONT_SIZE_PATTERN.search(style):
        style = FONT_SIZE_PATTERN.sub(f"font-size:{font_size}px", style)
    else:
        if style and not style.endswith(";"):
            style += ";"
        style += f"font-size:{font_size}px"
    element.setAttribute("style", style)
    element.setAttribute("font-size", _format_float(font_size))


def _extract_font_family(element: Element) -> str:
    style = element.getAttribute("style") or ""
    match = FONT_FAMILY_PATTERN.search(style)
    if match:
        return match.group(1).strip().strip("\"'")
    return ""


def _parse_length(value: str) -> Optional[float]:
    if not value:
        return None
    stripped = value.strip()
    if stripped.lower().endswith("px"):
        stripped = stripped[:-2]
    try:
        return float(stripped)
    except ValueError:
        return None


class _TransformState(NamedTuple):
    dx: float
    dy: float
    remaining: Optional[str]
    original: Optional[str]


def _parse_translate_arguments(argument: str) -> Tuple[float, float]:
    cleaned = argument.replace(",", " ").strip()
    if not cleaned:
        return 0.0, 0.0

    tokens = [token for token in cleaned.split() if token]
    dx = _parse_length(tokens[0]) if tokens else None
    dy = _parse_length(tokens[1]) if len(tokens) > 1 else None

    return (dx if dx is not None else 0.0), (dy if dy is not None else 0.0)


def _capture_transform_state(element: Element) -> _TransformState:
    if not element.hasAttribute("transform"):
        return _TransformState(0.0, 0.0, None, None)

    original = element.getAttribute("transform") or ""
    dx_total = 0.0
    dy_total = 0.0
    segments = []
    last_index = 0

    for match in _TRANSLATE_RE.finditer(original):
        segments.append(original[last_index : match.start()])
        offset_dx, offset_dy = _parse_translate_arguments(match.group(1))
        dx_total += offset_dx
        dy_total += offset_dy
        last_index = match.end()

    segments.append(original[last_index:])

    cleaned = " ".join(segment.strip() for segment in segments if segment.strip())
    remaining = cleaned or None

    return _TransformState(dx_total, dy_total, remaining, original)


def _offset_coordinate_string(value: str, delta: float) -> Optional[str]:
    tokens = [token for token in _COORDINATE_SPLIT_RE.split(value.strip()) if token]
    if not tokens:
        return None

    updated_tokens = []
    parsed_any = False
    for token in tokens:
        parsed = _parse_length(token)
        if parsed is None:
            updated_tokens.append(token)
            continue
        parsed_any = True
        suffix = "px" if token.lower().endswith("px") else ""
        updated_tokens.append(f"{_format_float(parsed + delta)}{suffix}")

    if not parsed_any:
        return None

    return " ".join(updated_tokens)


def _apply_coordinate_offset(element: Element, attribute: str, delta: float) -> None:
    if math.isclose(delta, 0.0, abs_tol=1e-9):
        return

    if element.hasAttribute(attribute):
        raw_value = element.getAttribute(attribute)
        updated = _offset_coordinate_string(raw_value, delta)
        if updated is not None:
            element.setAttribute(attribute, updated)
            return

    element.setAttribute(attribute, _format_float(delta))


def _extract_font_size(element: Element) -> Optional[float]:
    style = element.getAttribute("style") or ""
    match = FONT_SIZE_PATTERN.search(style)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            pass

    if element.hasAttribute("font-size"):
        length = _parse_length(element.getAttribute("font-size"))
        if length is not None:
            return length

    return None


def _adjust_font_size(element: Element, text_length: int, max_characters: Optional[int], reduction: float) -> None:
    if max_characters is None or reduction <= 0:
        return
    overflow = text_length - max_characters
    if overflow <= 0:
        return

    style = element.getAttribute("style")
    match = FONT_SIZE_PATTERN.search(style)
    if match:
        try:
            base_size = float(match.group(1))
        except ValueError:
            base_size = None
    else:
        base_size = None

    if base_size is None:
        # Default to a conservative font size if not specified.
        base_size = 12.0

    new_size = max(base_size - reduction * overflow, MIN_FONT_SIZE)
    if match:
        style = FONT_SIZE_PATTERN.sub(f"font-size:{new_size}px", style)
    else:
        if style and not style.endswith(";"):
            style += ";"
        style += f"font-size:{new_size}px"
    element.setAttribute("style", style)


def _measure_text_width(font: ImageFont.FreeTypeFont, text: str) -> float:
    if not text:
        return 0.0
    left, _, right, _ = font.getbbox(text)
    return float(right - left)


def _resolve_font_path(element: Element) -> Path:
    font_family = _extract_font_family(element)
    base_path = Path(__file__).resolve().parent
    if font_family and "marvin" in font_family.lower():
        candidate = base_path / "Marvin.ttf"
        if candidate.exists():
            return candidate
    candidate = base_path / "PlaypenSans-Medium.ttf"
    return candidate


def _extract_rect_width(rect: Element) -> Optional[float]:
    if rect.tagName.lower() != "rect":
        return None
    width = _parse_length(rect.getAttribute("width"))
    if width is None or width <= 0:
        return None
    return width


def _find_nearest_rect_width(element: Element) -> Optional[float]:
    """Attempt to infer the available width from a sibling rectangle."""

    def _iter_siblings(start: Optional[Node], step: str) -> Iterator[Element]:
        current = getattr(start, step)
        while current is not None:
            if current.nodeType == Node.ELEMENT_NODE:
                yield current  # type: ignore[misc]
            current = getattr(current, step)

    for sibling in _iter_siblings(element, "previousSibling"):
        width = _extract_rect_width(sibling)
        if width is not None:
            return width

    for sibling in _iter_siblings(element, "nextSibling"):
        width = _extract_rect_width(sibling)
        if width is not None:
            return width

    parent = element.parentNode
    while parent is not None and parent.nodeType != Node.ELEMENT_NODE:
        parent = parent.parentNode

    if isinstance(parent, Element):
        for node in parent.childNodes:
            if node.nodeType == Node.ELEMENT_NODE:
                width = _extract_rect_width(node)  # type: ignore[arg-type]
                if width is not None:
                    return width

    return None


def _compute_max_text_width(element: Element, template_lines: Sequence[str]) -> Optional[float]:
    font_size = _extract_font_size(element)
    if font_size is None:
        return None

    font_path = _resolve_font_path(element)
    if not font_path.exists():
        return None

    try:
        font = ImageFont.truetype(str(font_path), max(1, int(round(font_size))))
    except OSError:
        return None

    cleaned_lines = [line.strip() for line in template_lines if line.strip()]
    if not cleaned_lines:
        template_width = None
    else:
        template_width = max(
            (_measure_text_width(font, line) for line in cleaned_lines),
            default=None,
        )

    rect_width = _find_nearest_rect_width(element)

    if template_width is not None and rect_width is not None:
        return max(template_width, rect_width)
    if template_width is not None:
        return template_width
    return rect_width


class _WidthFitResult(NamedTuple):
    fits: bool
    max_width: Optional[float]
    baseline_y: Optional[float]
    font_size: float
    measured_width: Optional[float]
    font_path: Optional[Path]
    template_font_size: Optional[float]
    transform_dx: float
    transform_dy: float


def _split_text_into_two_lines(text: str) -> Sequence[str]:
    cleaned = text.strip()
    if not cleaned:
        return [""]

    words = cleaned.split()
    if len(words) <= 1:
        return [cleaned]

    best_index = 1
    best_diff = float("inf")
    for index in range(1, len(words)):
        first_line = " ".join(words[:index])
        second_line = " ".join(words[index:])
        diff = abs(len(first_line) - len(second_line))
        if diff < best_diff:
            best_diff = diff
            best_index = index

    first = " ".join(words[:best_index]).strip()
    second = " ".join(words[best_index:]).strip()
    return [line for line in (first, second) if line]


def _apply_alignment(element: Element, alignment: str) -> None:
    """Apply SVG text alignment while accepting common synonyms."""

    normalized = alignment.strip().lower()
    if normalized in {"center", "centre", "middle"}:
        element.setAttribute("text-anchor", "middle")
    elif normalized in {"right", "end"}:
        element.setAttribute("text-anchor", "end")
    elif normalized in {"left", "start"}:
        element.setAttribute("text-anchor", "start")


def _fit_text_within_width(
    element: Element,
    text: str,
    *,
    max_width: Optional[float],
    baseline_y: Optional[float],
    min_font_size: float = MIN_FONT_SIZE,
    alignment: str = "center",
    template_font_size: Optional[float] = None,
) -> _WidthFitResult:
    transform_state = _capture_transform_state(element)
    font_size = _extract_font_size(element)
    if font_size is None:
        font_size = template_font_size
    if font_size is None:
        font_size = 38.0

    current_size = max(float(font_size), min_font_size)

    font_path = _resolve_font_path(element)
    measured_width: Optional[float] = None
    font_for_measurement: Optional[ImageFont.FreeTypeFont] = None

    initial_measured_width: Optional[float] = None

    if font_path.exists():
        try:
            font_for_measurement = ImageFont.truetype(
                str(font_path), max(1, int(round(current_size)))
            )
        except OSError:
            font_for_measurement = None

    if font_for_measurement is not None:
        measured_width = _measure_text_width(font_for_measurement, text)
        initial_measured_width = measured_width
        while (
            max_width is not None
            and max_width > 0
            and measured_width is not None
            and measured_width > max_width
            and current_size > min_font_size
        ):
            new_size = max(current_size - 0.2, min_font_size)
            if math.isclose(new_size, current_size, rel_tol=1e-3, abs_tol=1e-3):
                break
            current_size = new_size
            try:
                font_for_measurement = ImageFont.truetype(
                    str(font_path), max(1, int(round(current_size)))
                )
            except OSError:
                font_for_measurement = None
                measured_width = None
                break
            measured_width = _measure_text_width(font_for_measurement, text)

    _set_font_size(element, round(current_size, 2))

    adjusted_baseline = (
        (baseline_y + transform_state.dy)
        if baseline_y is not None
        else None
    )

    if adjusted_baseline is not None:
        element.setAttribute("y", _format_float(adjusted_baseline))
    elif not math.isclose(transform_state.dy, 0.0, abs_tol=1e-9):
        _apply_coordinate_offset(element, "y", transform_state.dy)

    if not math.isclose(transform_state.dx, 0.0, abs_tol=1e-9):
        _apply_coordinate_offset(element, "x", transform_state.dx)

    element.setAttribute("dominant-baseline", "alphabetic")

    if alignment:
        _apply_alignment(element, alignment)

    if transform_state.remaining is not None:
        element.setAttribute("transform", transform_state.remaining)
    elif transform_state.original is not None and element.hasAttribute("transform"):
        element.removeAttribute("transform")

    width_tolerance = 0.05
    fits = True
    if (
        max_width is not None
        and max_width > 0
        and initial_measured_width is not None
        and initial_measured_width > max_width + width_tolerance
    ):
        fits = False
    if (
        max_width is not None
        and max_width > 0
        and measured_width is not None
        and measured_width > max_width + width_tolerance
    ):
        fits = False

    available_font_path = font_path if font_path.exists() else None

    final_baseline: Optional[float]
    if element.hasAttribute("y"):
        final_baseline = _parse_length(element.getAttribute("y"))
    else:
        final_baseline = adjusted_baseline

    return _WidthFitResult(
        fits,
        max_width,
        final_baseline,
        current_size,
        measured_width,
        available_font_path,
        template_font_size,
        transform_state.dx,
        transform_state.dy,
    )


def _apply_multiline_layout(
    element: Element,
    lines: Sequence[str],
    fit_result: _WidthFitResult,
    *,
    alignment: str = "center",
    min_font_size: float = MIN_FONT_SIZE,
    initial_size: Optional[float] = None,
) -> Tuple[float, Optional[float]]:
    if initial_size is not None:
        effective_size = max(initial_size, min_font_size)
    else:
        base_size = (
            fit_result.template_font_size
            if fit_result.template_font_size is not None
            else fit_result.font_size
        )
        effective_size = max(base_size, min_font_size)

    font_path = fit_result.font_path or _resolve_font_path(element)
    measured_width: Optional[float] = None
    font_for_metrics: Optional[ImageFont.FreeTypeFont] = None

    if font_path and font_path.exists():
        try:
            font_for_metrics = ImageFont.truetype(
                str(font_path), max(1, int(round(effective_size)))
            )
        except OSError:
            font_for_metrics = None

    if font_for_metrics is not None:
        measured_width = max(
            (_measure_text_width(font_for_metrics, line) for line in lines),
            default=0.0,
        )
        while (
            fit_result.max_width is not None
            and fit_result.max_width > 0
            and measured_width is not None
            and measured_width > fit_result.max_width
            and effective_size > min_font_size
        ):
            new_size = max(effective_size - 0.2, min_font_size)
            if math.isclose(new_size, effective_size, rel_tol=1e-3, abs_tol=1e-3):
                break
            effective_size = new_size
            try:
                font_for_metrics = ImageFont.truetype(
                    str(font_path), max(1, int(round(effective_size)))
                )
            except OSError:
                font_for_metrics = None
                measured_width = None
                break
            measured_width = max(
                (_measure_text_width(font_for_metrics, line) for line in lines),
                default=0.0,
            )

    effective_size = max(effective_size, min_font_size)
    _set_font_size(element, round(effective_size, 2))

    if alignment:
        _apply_alignment(element, alignment)

    baseline_reference = fit_result.baseline_y
    if baseline_reference is None and element.hasAttribute("y"):
        baseline_reference = _parse_length(element.getAttribute("y"))
    if baseline_reference is None:
        baseline_reference = 0.0

    baseline_spacing: float
    baseline_candidate: Optional[float] = None
    if font_for_metrics is not None:
        try:
            ascent, descent = font_for_metrics.getmetrics()
        except (AttributeError, OSError):
            ascent = descent = None
        if ascent is not None and descent is not None:
            baseline_candidate = float(ascent + descent)
        elif hasattr(font_for_metrics, "size"):
            baseline_candidate = float(getattr(font_for_metrics, "size"))

    if baseline_candidate is None or baseline_candidate <= 0:
        baseline_candidate = effective_size

    baseline_spacing = baseline_candidate * BASELINE_SPACING_SCALE
    if baseline_spacing <= 0:
        baseline_spacing = max(effective_size * BASELINE_SPACING_SCALE, 1.0)

    line_count = max(len(lines), 1)
    total_span = baseline_spacing * max(line_count - 1, 0)
    first_baseline = baseline_reference - (total_span / 2.0)

    element.setAttribute("y", _format_float(first_baseline))
    element.setAttribute("dominant-baseline", "alphabetic")

    _set_multiline_text(element, lines, line_height=baseline_spacing)

    return effective_size, measured_width


def _apply_two_line_layout(
    element: Element,
    text: str,
    fit_result: _WidthFitResult,
    *,
    alignment: str = "center",
    min_font_size: float = MIN_FONT_SIZE,
) -> Optional[Tuple[Sequence[str], float, Optional[float]]]:
    lines = _split_text_into_two_lines(text)
    if len(lines) < 2:
        return None

    effective_size, measured_width = _apply_multiline_layout(
        element,
        lines,
        fit_result,
        alignment=alignment,
        min_font_size=min_font_size,
    )

    return lines, effective_size, measured_width


def _shrink_two_line_text(
    element: Element,
    lines: Sequence[str],
    fit_result: _WidthFitResult,
    *,
    alignment: str,
    absolute_min_font_size: float,
) -> Optional[Tuple[float, Optional[float]]]:
    max_width = fit_result.max_width
    if max_width is None or max_width <= 0:
        return None

    current_size = _extract_font_size(element)
    if current_size is None:
        current_size = fit_result.font_size
    if current_size is None:
        current_size = absolute_min_font_size

    measured_width: Optional[float] = None
    applied_size = current_size

    for _ in range(50):
        applied_size, measured_width = _apply_multiline_layout(
            element,
            lines,
            fit_result,
            alignment=alignment,
            min_font_size=absolute_min_font_size,
            initial_size=applied_size,
        )
        if measured_width is None or measured_width <= max_width:
            break
        if applied_size <= absolute_min_font_size or math.isclose(
            applied_size, absolute_min_font_size, abs_tol=1e-3
        ):
            break
        applied_size = max(applied_size - 0.5, absolute_min_font_size)

    return applied_size, measured_width


def _update_text_group(group: Element, text: str, *, max_characters: Optional[int] = None, reduction: float = 0.0) -> None:
    text_elements = list(group.getElementsByTagName("text"))

    base_alignment = _resolve_layer_alignment(group)
    if base_alignment is None:
        group_id = group.getAttribute("id").lower() if group.hasAttribute("id") else ""
        if group_id in LEFT_ALIGNED_GROUPS:
            base_alignment = "left"
        elif group_id in CENTER_ALIGNED_GROUPS:
            base_alignment = "center"

    cached_dimensions = []
    for text_element in text_elements:
        template_lines = _extract_template_lines(text_element)
        max_width = _compute_max_text_width(text_element, template_lines)
        baseline_y = _parse_length(text_element.getAttribute("y") if text_element.hasAttribute("y") else "")
        template_font_size = _extract_font_size(text_element)
        if text_element.hasAttribute("x"):
            original_x = text_element.getAttribute("x")
            has_x = True
        else:
            original_x = None
            has_x = False
        if text_element.hasAttribute("text-anchor"):
            original_anchor = text_element.getAttribute("text-anchor")
            has_anchor = True
        else:
            original_anchor = None
            has_anchor = False
        cached_dimensions.append(
            {
                "max_width": max_width,
                "baseline_y": baseline_y,
                "template_font_size": template_font_size,
                "original_x": original_x,
                "has_x": has_x,
                "original_anchor": original_anchor,
                "has_anchor": has_anchor,
            }
        )

    for index, text_element in enumerate(text_elements):
        element_alignment = _resolve_layer_alignment(text_element) or base_alignment

        _set_text(text_element, text)
        _adjust_font_size(text_element, len(text), max_characters, reduction)

        cache = cached_dimensions[index] if index < len(cached_dimensions) else {}
        max_width = cache.get("max_width")
        baseline_y = cache.get("baseline_y")
        template_font_size = cache.get("template_font_size")
        original_x = cache.get("original_x")
        has_original_x = cache.get("has_x", False)
        original_anchor = cache.get("original_anchor")
        has_original_anchor = cache.get("has_anchor", False)

        fit_alignment = element_alignment or "center"
        fit_result = _fit_text_within_width(
            text_element,
            text,
            max_width=max_width,
            baseline_y=baseline_y,
            alignment=fit_alignment,
            template_font_size=template_font_size,
        )

        if element_alignment:
            _apply_alignment(text_element, element_alignment)

        multiline_applied = False
        if (
            text.strip()
            and fit_result is not None
            and not fit_result.fits
            and fit_result.max_width is not None
        ):
            fallback_alignment = element_alignment or "center"
            fallback_result = _apply_two_line_layout(
                text_element,
                text,
                fit_result,
                alignment=fallback_alignment,
            )

            if fallback_result is not None:
                multiline_applied = True
                lines, _size, measured_width = fallback_result
                if (
                    fit_result.max_width is not None
                    and fit_result.max_width > 0
                    and measured_width is not None
                    and measured_width > fit_result.max_width
                ):
                    _shrink_two_line_text(
                        text_element,
                        lines,
                        fit_result,
                        alignment=fallback_alignment,
                        absolute_min_font_size=MULTILINE_MIN_FONT_SIZE,
                    )

        if (
            not multiline_applied
            and fit_result.fits
            and has_original_x
            and original_x is not None
            and math.isclose(fit_result.transform_dx, 0.0, abs_tol=1e-9)
        ):
            text_element.setAttribute("x", original_x)

        if not multiline_applied and has_original_anchor:
            if original_anchor:
                text_element.setAttribute("text-anchor", original_anchor)
            elif text_element.hasAttribute("text-anchor"):
                text_element.removeAttribute("text-anchor")


def _update_address_group(group: Element, text: str) -> None:
    lines = [line.strip() for line in text.replace("\r", "").split("\n") if line.strip()]
    if not lines:
        lines = [""]

    base_alignment = _resolve_layer_alignment(group) or "left"

    for text_element in group.getElementsByTagName("text"):
        element_alignment = _resolve_layer_alignment(text_element) or base_alignment
        if element_alignment:
            _apply_alignment(text_element, element_alignment)
        _set_multiline_text(text_element, lines)


def _find_group_map(doc: Document) -> Dict[str, Element]:
    groups = {}
    for group in doc.getElementsByTagName("g"):
        group_id = group.getAttribute("id").lower() if group.hasAttribute("id") else ""
        if group_id:
            groups[group_id] = group
    return groups


def _find_template_file(directory: Path, stem: str) -> Optional[Path]:
    for extension in (".svg", ".SVG"):
        candidate = directory / f"{stem}{extension}"
        if candidate.exists():
            return candidate
    return None


def _guardian_type(value: object) -> Optional[int]:
    if _is_missing(value):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        value_str = str(value).strip()
        if value_str.isdigit():
            return int(value_str)
    return None


def _copy_photo(source: Path, destination_dir: Path) -> Optional[str]:
    if not source.exists():
        return None
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / source.name
    shutil.copyfile(source, destination)
    try:
        relative_path = destination.relative_to(destination_dir.parent)
    except ValueError:
        return destination.name
    return relative_path.as_posix()


def _prepare_working_directory(template_dir: Path, working_dir: Path) -> None:
    if working_dir.exists():
        shutil.rmtree(working_dir)
    shutil.copytree(template_dir, working_dir)


def _process_svg(svg_path: Path, updates: Dict[str, Tuple[str, Optional[int], float]], image_updates: Dict[str, str], address_updates: Dict[str, str]) -> bool:
    if svg_path is None or not svg_path.exists():
        return False

    doc = parse(str(svg_path))
    group_map = _find_group_map(doc)

    for group_id, (text, max_chars, reduction) in updates.items():
        group = group_map.get(group_id)
        if group is None:
            continue
        _update_text_group(group, text, max_characters=max_chars, reduction=reduction)

    for group_id, text in address_updates.items():
        group = group_map.get(group_id)
        if group is None:
            continue
        _update_address_group(group, text)

    for group_id, image_name in image_updates.items():
        if not image_name:
            continue
        group = group_map.get(group_id)
        if group is None:
            continue
        for image in group.getElementsByTagName("image"):
            image.setAttribute("xlink:href", image_name)

    with open(svg_path, "w", encoding="utf-8") as handle:
        handle.write(doc.toxml())
    return True


def _format_date(value: str) -> str:
    parts = value.split("-")
    if len(parts) == 3:
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    return value


def personalize_id_card(
    record: Dict[str, object],
    *,
    template_root: Path = DEFAULT_TEMPLATE_ROOT,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
    photo_root: Path = DEFAULT_PHOTO_ROOT,
) -> bool:
    school_name_raw = _normalise_string(record.get("school_name"))
    if not school_name_raw:
        return False

    school_id = _normalise_string(record.get("school_id"))
    if not school_id:
        return False

    user_id = _normalise_string(record.get("user_id"))
    if not user_id:
        return False

    template_dir = template_root / school_id
    if not template_dir.exists():
        raise TemplateNotFoundError(f"Template directory not found for school: {school_id}")

    front_template = _find_template_file(template_dir, "FRONT")
    back_template = _find_template_file(template_dir, "BACK")
    if front_template is None and back_template is None:
        raise TemplateNotFoundError(f"No SVG templates found for school: {school_name_raw}")

    first_name = _normalise_string(record.get("first_name"))
    last_name = _normalise_string(record.get("last_name"))
    child_output_base = _build_child_output_base(first_name, last_name, school_name_raw)

    school_output_dir = output_root / school_id
    child_output_dir = school_output_dir / child_output_base
    photos_output_dir = child_output_dir / "working" / "images"

    _ensure_directory(child_output_dir)
    working_dir = child_output_dir / "working"
    _prepare_working_directory(template_dir, working_dir)

    class_name = _normalise_string(record.get("class_name"))
    blood_group = _normalise_string(record.get("blood_group"))
    age = _normalise_string(record.get("age"))
    date_of_birth = _normalise_string(record.get("date_of_birth"))
    if date_of_birth:
        date_of_birth = _format_date(date_of_birth)

    student_id = _normalise_string(record.get("student_id"))
    admission_number = _normalise_string(record.get("admission_number"))
    roll_number = _normalise_string(record.get("roll_number"))
    register_number = _normalise_string(record.get("register_number"))
    date_of_issued = _normalise_string(record.get("date_of_issued"))
    if date_of_issued:
        date_of_issued = _format_date(date_of_issued)

    expiry_date = _normalise_string(record.get("expiry_date"))
    if expiry_date:
        expiry_date = _format_date(expiry_date)

    gender = _normalise_string(record.get("gender"))
    current_address = _normalise_string(record.get("current_address"))

    department_value = record.get("department")
    department = (
        custom_title_case(_normalise_string(department_value))
        if not _is_missing(department_value)
        else ""
    )
    employee_id = _normalise_string(record.get("employee_id"))
    email = _normalise_string(record.get("email"))

    guardian_1_type = _guardian_type(record.get("guardian_1_type"))
    guardian_2_type = _guardian_type(record.get("guardian_2_type"))

    guardian_1_name = custom_title_case(_normalise_string(record.get("guardian_1_name"))) if not _is_missing(record.get("guardian_1_name")) else ""
    guardian_2_name = custom_title_case(_normalise_string(record.get("guardian_2_name"))) if not _is_missing(record.get("guardian_2_name")) else ""

    guardian_1_mobile = _normalise_string(record.get("guardian_1_mobile"))
    guardian_2_mobile = _normalise_string(record.get("guardian_2_mobile"))

    guardian_1_id = _normalise_string(record.get("guardian_1_id"))
    guardian_2_id = _normalise_string(record.get("guardian_2_id"))

    father_name = ""
    mother_name = ""
    father_contact = ""
    mother_contact = ""
    father_photo_id = ""
    mother_photo_id = ""

    if guardian_1_type == 0:
        father_name = guardian_1_name
        father_contact = guardian_1_mobile
        father_photo_id = guardian_1_id
    elif guardian_1_type == 1:
        mother_name = guardian_1_name
        mother_contact = guardian_1_mobile
        mother_photo_id = guardian_1_id

    if guardian_2_type == 0:
        father_name = guardian_2_name or father_name
        father_contact = guardian_2_mobile or father_contact
        father_photo_id = guardian_2_id or father_photo_id
    elif guardian_2_type == 1:
        mother_name = guardian_2_name or mother_name
        mother_contact = guardian_2_mobile or mother_contact
        mother_photo_id = guardian_2_id or mother_photo_id

    full_name_parts = [part for part in (first_name, last_name) if part]
    full_name = custom_title_case(" ".join(full_name_parts)).strip()

    school_branch = clean_branch_name(school_name_raw)

    school_id = _normalise_string(record.get("school_id"))
    photo_school_root = photo_root / school_id if school_id else photo_root

    child_photo_path = photo_school_root / "PARTIAL" / f"{user_id}.png"
    father_photo_path = photo_school_root / "PARTIAL" / f"{father_photo_id}.png" if father_photo_id else None
    mother_photo_path = photo_school_root / "PARTIAL" / f"{mother_photo_id}.png" if mother_photo_id else None

    child_photo_name = _copy_photo(child_photo_path, photos_output_dir) if child_photo_path else None
    father_photo_name = _copy_photo(father_photo_path, photos_output_dir) if father_photo_path else None
    mother_photo_name = _copy_photo(mother_photo_path, photos_output_dir) if mother_photo_path else None

    text_updates_front = {
        "name": (full_name, 15, 0.6),
        "grade": (class_name, 14, 0.5),
        "branch": (school_branch, 20, 0.5),
        "mcontact": (mother_contact, 14, 0.5),
        "fcontact": (father_contact, 14, 0.5),
        "fname": (father_name, 30, 0.3),
        "mname": (mother_name, 30, 0.3),
        "blood": (blood_group, 50, 0.5),
        "age": (age, 50, 0.5),
        "dob": (date_of_birth, 12, 0.5),
        "gender": (gender, 500, 0.5),
    }

    text_updates_back = {
        "name": (full_name, 15, 0.6),
        "mcontact": (mother_contact, 15, 0.5),
        "fcontact": (father_contact, 15, 0.5),
        "fname": (father_name, 30, 0.3),
        "mname": (mother_name, 30, 0.3),
        "blood": (blood_group, 5, 0.5),
        "dob": (date_of_birth, 12, 0.5),
    }

    additional_text_updates = {
        "student_id": (student_id, 20, 0.5),
        "admission_number": (admission_number, 20, 0.5),
        "roll_number": (roll_number, 20, 0.5),
        "register_number": (register_number, 20, 0.5),
        "date_of_issued": (date_of_issued, 12, 0.5),
        "expiry_date": (expiry_date, 12, 0.5),
        "department": (department, 30, 0.3),
        "employee_id": (employee_id, 20, 0.5),
        "email": (email, 40, 0.5),
    }

    text_updates_front.update(additional_text_updates)
    text_updates_back.update(additional_text_updates)

    text_updates_front["blood_group"] = text_updates_front["blood"]
    text_updates_back["blood_group"] = text_updates_back["blood"]

    address_updates_front = {
        "address": current_address,
        "address2": current_address,
    }
    address_updates_back = {"address": current_address}

    image_updates_front = {"pic1": child_photo_name or ""}
    image_updates_back = {
        "pic2": mother_photo_name or "",
        "pic3": father_photo_name or "",
    }

    generated = False

    if front_template is not None:
        front_svg_path = working_dir / front_template.name
        _process_svg(front_svg_path, text_updates_front, image_updates_front, address_updates_front)
        front_pdf_path = child_output_dir / f"{child_output_base}_FRONT.pdf"
        callInkscape(str(front_svg_path), str(front_pdf_path))
        generated = True

    if back_template is not None:
        back_svg_path = working_dir / back_template.name
        _process_svg(back_svg_path, text_updates_back, image_updates_back, address_updates_back)
        back_pdf_path = child_output_dir / f"{child_output_base}_BACK.pdf"
        callInkscape(str(back_svg_path), str(back_pdf_path))
        generated = True

    return generated


def dc_sanitize(value: str) -> str:
    from dc import _sanitize_for_path

    return _sanitize_for_path(value, "School")


def generate_id_cards(
    records: Iterable[Dict[str, object]],
    *,
    template_root: Path = DEFAULT_TEMPLATE_ROOT,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
    photo_root: Path = DEFAULT_PHOTO_ROOT,
) -> int:
    count = 0
    for record in records:
        try:
            if personalize_id_card(
                record,
                template_root=template_root,
                output_root=output_root,
                photo_root=photo_root,
            ):
                count += 1
        except TemplateNotFoundError as exc:
            print(exc)
            continue
    return count


def load_records_from_csv(csv_path: Path) -> Iterator[Dict[str, str]]:
    with csv_path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            yield row


def generate_id_cards_from_csv(
    csv_path: Path,
    *,
    template_root: Path = DEFAULT_TEMPLATE_ROOT,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
    photo_root: Path = DEFAULT_PHOTO_ROOT,
) -> int:
    return generate_id_cards(
        load_records_from_csv(csv_path),
        template_root=template_root,
        output_root=output_root,
        photo_root=photo_root,
    )


def _parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate ID cards from a CSV input sheet.")
    parser.add_argument("csv_path", type=Path, help="Path to the CSV file containing ID card data")
    parser.add_argument(
        "--template-root",
        type=Path,
        default=DEFAULT_TEMPLATE_ROOT,
        help="Directory containing per-school SVG templates",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=DEFAULT_OUTPUT_ROOT,
        help="Directory where generated ID cards will be written",
    )
    parser.add_argument(
        "--photo-root",
        type=Path,
        default=DEFAULT_PHOTO_ROOT,
        help="Directory containing student and guardian photos",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = _parse_args(argv)
    count = generate_id_cards_from_csv(
        args.csv_path,
        template_root=args.template_root,
        output_root=args.output_root,
        photo_root=args.photo_root,
    )
    print(f"Generated {count} ID card(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
