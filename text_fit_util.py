"""Utility helpers for fitting SVG text within bounding rectangles."""
from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Dict, List, Optional

from PIL import ImageFont
from xml.dom.minidom import Element

_FONT_SIZE_PATTERN = re.compile(r"font-size\s*:\s*([0-9.]+)px")
_FONT_FAMILY_PATTERN = re.compile(r"font-family\s*:\s*([^;]+)")

_WIDTH_PADDING_RATIO = 0.95
_HEIGHT_PADDING_RATIO = 0.95
_LINE_SPACING_EM = 1.1
_MIN_SCALE = 0.7
_MAX_SCALE = 1.1

_DEFAULT_FONT = Path(__file__).resolve().parent / "PlaypenSans-Medium.ttf"
_MARVIN_FONT = Path(__file__).resolve().parent / "Marvin.ttf"


def _format_float(value: float) -> str:
    """Format floats for XML attributes without trailing zeros."""

    text = f"{value:.4f}".rstrip("0").rstrip(".")
    return text or "0"


def _parse_float(value: Optional[str], default: Optional[float] = None) -> Optional[float]:
    if value is None:
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _extract_font_size(element: Element, default: float = 38.0) -> float:
    style = element.getAttribute("style") if element.hasAttribute("style") else ""
    match = _FONT_SIZE_PATTERN.search(style)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            pass
    return default


def _set_font_size(element: Element, size: float) -> None:
    style = element.getAttribute("style") if element.hasAttribute("style") else ""
    formatted = f"{size:.2f}".rstrip("0").rstrip(".")
    replacement = f"font-size:{formatted}px"
    if _FONT_SIZE_PATTERN.search(style):
        style = _FONT_SIZE_PATTERN.sub(replacement, style)
    else:
        if style and not style.strip().endswith(";"):
            style = style.strip() + ";"
        style += replacement
    element.setAttribute("style", style)


def _extract_font_family(element: Element) -> str:
    if element.hasAttribute("font-family"):
        value = element.getAttribute("font-family").strip()
        if value:
            return value
    style = element.getAttribute("style") if element.hasAttribute("style") else ""
    match = _FONT_FAMILY_PATTERN.search(style)
    if match:
        return match.group(1).strip()
    return ""


def _resolve_font_path(element: Element, font_path: Optional[Path] = None) -> Optional[Path]:
    candidates: List[Path] = []
    if font_path is not None:
        candidates.append(Path(font_path))

    font_family = _extract_font_family(element).lower()
    if "marvin" in font_family:
        candidates.append(_MARVIN_FONT)

    candidates.append(_DEFAULT_FONT)

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0] if candidates else None


def _load_font(font_path: Path, size: float) -> Optional[ImageFont.FreeTypeFont]:
    try:
        effective_size = max(1, int(round(size)))
        return ImageFont.truetype(str(font_path), effective_size)
    except (OSError, TypeError):
        return None


def _measure_text_width(font: ImageFont.FreeTypeFont, text: str) -> float:
    if not text:
        return 0.0
    try:
        return float(font.getlength(text))
    except AttributeError:
        left, _, right, _ = font.getbbox(text)
        return float(right - left)


def _split_text_into_two_lines(text: str) -> List[str]:
    cleaned = text.strip()
    if not cleaned:
        return ["", ""]

    words = cleaned.split()
    if len(words) <= 1:
        midpoint = max(1, len(cleaned) // 2)
        return [cleaned[:midpoint].strip(), cleaned[midpoint:].strip() or ""]

    best_index = 1
    best_diff = float("inf")
    for index in range(1, len(words)):
        first = " ".join(words[:index]).strip()
        second = " ".join(words[index:]).strip()
        diff = abs(len(first) - len(second))
        if diff < best_diff:
            best_diff = diff
            best_index = index

    first_line = " ".join(words[:best_index]).strip()
    second_line = " ".join(words[best_index:]).strip()
    if not second_line:
        midpoint = max(1, len(first_line) // 2)
        second_line = first_line[midpoint:].strip()
        first_line = first_line[:midpoint].strip()
    return [first_line, second_line or ""]


def _truncate_line_to_width(
    font: ImageFont.FreeTypeFont, line: str, target_width: float
) -> str:
    trimmed = line.rstrip()
    if not trimmed:
        return trimmed

    if _measure_text_width(font, trimmed) <= target_width:
        return trimmed

    ellipsis = "\u2026"
    while trimmed and _measure_text_width(font, trimmed + ellipsis) > target_width:
        trimmed = trimmed[:-1].rstrip()

    return (trimmed + ellipsis) if trimmed else ellipsis


def _clear_children(element: Element) -> None:
    while element.firstChild is not None:
        element.removeChild(element.firstChild)


def _set_single_line_text(element: Element, text: str) -> None:
    _clear_children(element)
    element.appendChild(element.ownerDocument.createTextNode(text))


def _set_multiline_text(
    element: Element,
    lines: List[str],
    center_x: float,
    line_spacing: float,
) -> None:
    _clear_children(element)
    document = element.ownerDocument
    half_spacing = line_spacing / 2

    for index, line in enumerate(lines):
        tspan = document.createElement("tspan")
        tspan.setAttribute("x", _format_float(center_x))
        if index == 0:
            tspan.setAttribute("dy", f"-{half_spacing:.2f}em")
        else:
            tspan.setAttribute("dy", f"{line_spacing:.2f}em")
        tspan.appendChild(document.createTextNode(line))
        element.appendChild(tspan)


def _fit_text_within_rect(
    text_element: Element,
    rect_element: Optional[Element],
    text: str,
    *,
    font_path: Optional[Path] = None,
    width_padding: float = _WIDTH_PADDING_RATIO,
    min_scale: float = _MIN_SCALE,
    max_scale: float = _MAX_SCALE,
    line_spacing: float = _LINE_SPACING_EM,
) -> Dict[str, object]:
    """Fit ``text`` within the ``rect_element`` while updating ``text_element``.

    Returns diagnostic information about the applied sizing.
    """

    result: Dict[str, object] = {
        "font_size": None,
        "lines": [text],
        "was_split": False,
        "was_truncated": False,
    }

    if text_element is None or rect_element is None:
        return result

    rect_width = _parse_float(rect_element.getAttribute("width"))
    rect_height = _parse_float(rect_element.getAttribute("height"))
    rect_x = _parse_float(rect_element.getAttribute("x"), 0.0) or 0.0
    rect_y = _parse_float(rect_element.getAttribute("y"), 0.0) or 0.0

    if rect_width is None or rect_width <= 0:
        return result

    center_x = rect_x + rect_width / 2
    center_y = rect_y + (rect_height / 2 if rect_height else 0.0)

    text_element.setAttribute("text-anchor", "middle")
    text_element.setAttribute("dominant-baseline", "middle")
    text_element.setAttribute("x", _format_float(center_x))
    text_element.setAttribute("y", _format_float(center_y))
    if text_element.hasAttribute("transform"):
        text_element.removeAttribute("transform")

    base_size = _extract_font_size(text_element)
    if base_size <= 0:
        base_size = 1.0

    min_size = base_size * min_scale
    max_size = base_size * max_scale

    resolved_font_path = _resolve_font_path(text_element, font_path)
    if resolved_font_path is None:
        _set_font_size(text_element, base_size)
        _set_single_line_text(text_element, text)
        result["font_size"] = base_size
        return result

    current_size = min(max(base_size, min_size), max_size)
    font = _load_font(resolved_font_path, current_size)
    if font is None:
        _set_font_size(text_element, current_size)
        _set_single_line_text(text_element, text)
        result["font_size"] = current_size
        return result

    target_width = rect_width * width_padding
    text_width = _measure_text_width(font, text)

    if text_width <= target_width:
        _set_font_size(text_element, current_size)
        _set_single_line_text(text_element, text)
        result["font_size"] = current_size
        return result

    step = 0.5
    while text_width > target_width and current_size > min_size:
        new_size = max(current_size - step, min_size)
        if math.isclose(new_size, current_size, rel_tol=1e-3, abs_tol=1e-3):
            break
        current_size = new_size
        font = _load_font(resolved_font_path, current_size)
        if font is None:
            break
        text_width = _measure_text_width(font, text)

    if font is None:
        _set_font_size(text_element, current_size)
        _set_single_line_text(text_element, text)
        result["font_size"] = current_size
        return result

    if text_width <= target_width:
        _set_font_size(text_element, current_size)
        _set_single_line_text(text_element, text)
        result["font_size"] = current_size
        return result

    # Two-line fallback.
    lines = _split_text_into_two_lines(text)
    if len(lines) < 2:
        lines.append("")

    available_height = (rect_height or 0.0) * _HEIGHT_PADDING_RATIO if rect_height else None
    if available_height:
        height_units = (len(lines) - 1) * line_spacing + 1
        max_height_size = available_height / max(height_units, 1e-6)
    else:
        max_height_size = max_size

    max_allowed = min(max_height_size, max_size)
    if max_allowed <= 0:
        max_allowed = current_size

    if max_allowed < min_size:
        min_limit = max_allowed
    else:
        min_limit = min_size

    current_size = min(base_size, max_allowed)
    if current_size < min_limit:
        current_size = min_limit

    font = _load_font(resolved_font_path, current_size)
    if font is None:
        _set_font_size(text_element, current_size)
        _set_single_line_text(text_element, text)
        result["font_size"] = current_size
        return result

    line_widths = [_measure_text_width(font, line) for line in lines]
    max_line_width = max(line_widths) if line_widths else 0.0

    while max_line_width > target_width and current_size > min_limit:
        new_size = max(current_size - step, min_limit)
        if math.isclose(new_size, current_size, rel_tol=1e-3, abs_tol=1e-3):
            break
        current_size = new_size
        font = _load_font(resolved_font_path, current_size)
        if font is None:
            break
        line_widths = [_measure_text_width(font, line) for line in lines]
        max_line_width = max(line_widths) if line_widths else 0.0

    truncated = False
    if font is not None:
        adjusted_lines: List[str] = []
        for line in lines:
            updated_line = _truncate_line_to_width(font, line, target_width)
            if updated_line != line:
                truncated = True
            adjusted_lines.append(updated_line)
        lines = adjusted_lines

    _set_font_size(text_element, current_size)
    text_element.setAttribute("y", _format_float(center_y))
    _set_multiline_text(text_element, lines, center_x, line_spacing)

    result.update(
        {
            "font_size": current_size,
            "lines": lines,
            "was_split": True,
            "was_truncated": truncated,
        }
    )
    return result


__all__ = ["_fit_text_within_rect"]

