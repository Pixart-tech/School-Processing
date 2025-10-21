"""Utilities for generating personalised ID card PDFs from SVG templates."""
from __future__ import annotations
import argparse
import csv
import math
import re
import shutil
from pathlib import Path
from typing import Dict, Iterable, Iterator, Optional, Sequence, Tuple, Set
from PIL import ImageFont
from xml.dom.minidom import Document, Element, parse

from doc_maker import callInkscape


DEFAULT_TEMPLATE_ROOT = Path(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL ID CARD SRC")
DEFAULT_OUTPUT_ROOT = Path("ID Cards")
DEFAULT_PHOTO_ROOT = Path(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL_PHOTOS")
TEMP_ROOT = Path("Temp")
ID_CARD_WORKING_SUBDIR = "IDCardWorking"


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


def _extract_outer_code_prefix(value: object) -> Optional[str]:
    normalised = _normalise_string(value)
    if not normalised:
        return None
    digits = "".join(ch for ch in normalised if ch.isdigit())
    if len(digits) >= 3:
        return digits[:3]
    if len(normalised) >= 3:
        return normalised[:3]
    return None


def _normalise_directory_name(value: str) -> str:
    return re.sub(r"[^0-9a-z]+", "_", value.lower()).strip("_")


def _resolve_template_directory(
    template_root: Path, school_id: str, school_name: str
) -> Optional[Path]:
    normalised_school_id = _normalise_string(school_id)
    normalised_school_name = _normalise_string(school_name)

    if normalised_school_id:
        direct = template_root / normalised_school_id
        if direct.is_dir():
            return direct

    sanitized_school_name = (
        _sanitize_filename_component(normalised_school_name, "school")
        if normalised_school_name
        else ""
    )
    normalised_school_key = (
        _normalise_directory_name(sanitized_school_name)
        if sanitized_school_name
        else ""
    )

    school_name_tokens: Set[str] = set()
    if normalised_school_key:
        school_name_tokens = {
            token for token in normalised_school_key.split("_") if token
        }

    school_id_variants: Set[str] = set()
    if normalised_school_id:
        lower_id = normalised_school_id.lower()
        school_id_variants.add(lower_id)
        stripped = lower_id.lstrip("0")
        if stripped:
            school_id_variants.add(stripped)
        if lower_id.isdigit():
            school_id_variants.add(lower_id.zfill(3))

    best_match: Optional[Path] = None
    best_score = 0

    try:
        candidates = list(template_root.iterdir())
    except FileNotFoundError:
        candidates = []

    for entry in candidates:
        if not entry.is_dir():
            continue

        entry_key = _normalise_directory_name(entry.name)
        if not entry_key:
            continue

        entry_tokens = {token for token in entry_key.split("_") if token}

        score = 0

        if school_id_variants and (school_id_variants & entry_tokens):
            score += 2

        if normalised_school_key:
            if normalised_school_key in entry_key:
                score += 1
            elif school_name_tokens and (school_name_tokens & entry_tokens):
                score += 1

        if score > best_score:
            best_match = entry
            best_score = score

    if best_match is not None:
        return best_match

    if sanitized_school_name:
        fallback = template_root / sanitized_school_name
        if fallback.is_dir():
            return fallback

    return None


def _candidate_template_directories(
    template_root: Path, school_id: str, school_name: str
) -> Iterator[Path]:
    seen: Set[Path] = set()

    direct = template_root / _normalise_string(school_id)
    if direct.is_dir():
        seen.add(direct)
        yield direct

    resolved = _resolve_template_directory(template_root, school_id, school_name)
    if resolved is not None and resolved not in seen:
        seen.add(resolved)
        yield resolved

    if template_root not in seen:
        yield template_root


def _get_temp_working_dir(prefix: str, subdirectory: str) -> Path:
    base = TEMP_ROOT / prefix
    base.mkdir(parents=True, exist_ok=True)
    return base / subdirectory


def _working_file_path(template_file: Path, template_dir: Path, working_dir: Path) -> Path:
    try:
        relative_path = template_file.relative_to(template_dir)
    except ValueError:
        relative_path = Path(template_file.name)
    return working_dir / relative_path


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


def _set_text(element: Element, text: str) -> None:
    while element.firstChild:
        element.removeChild(element.firstChild)
    element.appendChild(element.ownerDocument.createTextNode(text))


def _set_font_size(element: Element, font_size: float) -> None:
    style = element.getAttribute("style") or ""
    if FONT_SIZE_PATTERN.search(style):
        style = FONT_SIZE_PATTERN.sub(f"font-size:{font_size}px", style)
    else:
        if style and not style.endswith(";"):
            style += ";"
        style += f"font-size:{font_size}px"
    element.setAttribute("style", style)


def _extract_font_family(element: Element) -> str:
    style = element.getAttribute("style") or ""
    match = FONT_FAMILY_PATTERN.search(style)
    if match:
        return match.group(1).strip().strip("\"'")
    return ""


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

    new_size = max(base_size - reduction * overflow, 1.0)
    if match:
        style = FONT_SIZE_PATTERN.sub(f"font-size:{new_size}px", style)
    else:
        if style and not style.endswith(";"):
            style += ";"
        style += f"font-size:{new_size}px"
    element.setAttribute("style", style)


def _fit_text_within_rect(group: Element, element: Element, text: str, index: int) -> None:
    rects = list(group.getElementsByTagName("rect"))
    if not rects:
        return

    rect = rects[index] if index < len(rects) else rects[0]
    width_attr = rect.getAttribute("width") if rect.hasAttribute("width") else ""
    try:
        rect_width = float(width_attr)
    except (TypeError, ValueError):
        return
    if rect_width <= 0:
        return

    font_size_match = FONT_SIZE_PATTERN.search(element.getAttribute("style") or "")
    if font_size_match:
        try:
            font_size = float(font_size_match.group(1))
        except ValueError:
            font_size = None
    else:
        font_size = None

    if font_size is None:
        font_size = 38.0

    font_family = _extract_font_family(element)
    font_path = Path(__file__).resolve().parent / "PlaypenSans-Medium.ttf"
    if "Marvin" in font_family:
        font_path = Path(__file__).resolve().parent / "Marvin.ttf"

    if not font_path.exists():
        return

    display_text = text.title()
    current_size = float(font_size)

    try:
        font = ImageFont.truetype(str(font_path), max(1, int(math.floor(current_size))))
    except OSError:
        return

    left, _, right, _ = font.getbbox(display_text)
    text_width = right - left

    while text_width > rect_width and current_size > 1:
        current_size -= 0.2
        try:
            font = ImageFont.truetype(str(font_path), max(1, int(math.floor(current_size))))
        except OSError:
            break
        left, _, right, _ = font.getbbox(display_text)
        text_width = right - left

    if current_size != font_size:
        _set_font_size(element, round(current_size, 2))


def _update_text_group(group: Element, text: str, *, max_characters: Optional[int] = None, reduction: float = 0.0) -> None:
    text_elements = list(group.getElementsByTagName("text"))
    for index, text_element in enumerate(text_elements):
        _set_text(text_element, text)
        _adjust_font_size(text_element, len(text), max_characters, reduction)
        if len(text) > 12:
            _fit_text_within_rect(group, text_element, text, index)


def _update_address_group(group: Element, text: str) -> None:
    lines = [line.strip() for line in text.replace("\r", "").split("\n") if line.strip()]
    if not lines:
        lines = [""]
    for text_element in group.getElementsByTagName("text"):
        while text_element.firstChild:
            text_element.removeChild(text_element.firstChild)

        document = text_element.ownerDocument
        base_x = text_element.getAttribute("x") if text_element.hasAttribute("x") else ""
        first_line = True

        for line in lines:
            if first_line:
                text_element.appendChild(document.createTextNode(line))
                first_line = False
                continue

            tspan = document.createElement("tspan")
            if base_x:
                tspan.setAttribute("x", base_x)
            tspan.setAttribute("dy", "1em")
            tspan.appendChild(document.createTextNode(line))
            text_element.appendChild(tspan)


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
    working_dir.parent.mkdir(parents=True, exist_ok=True)
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

    outer_prefix = _extract_outer_code_prefix(record.get("outer_code"))
    if not outer_prefix:
        return False

    template_group_dir = template_root / outer_prefix
    if not template_group_dir.is_dir():
        raise TemplateNotFoundError(
            f"Template directory not found for outer code prefix: {outer_prefix}"
        )

    template_dir: Optional[Path] = None
    front_template: Optional[Path] = None
    back_template: Optional[Path] = None

    for candidate in _candidate_template_directories(
        template_group_dir, school_id, school_name_raw
    ):
        front_candidate = _find_template_file(candidate, "FRONT")
        back_candidate = _find_template_file(candidate, "BACK")
        if front_candidate is None and back_candidate is None:
            continue
        template_dir = candidate
        front_template = front_candidate
        back_template = back_candidate
        break

    if template_dir is None:
        raise TemplateNotFoundError(
            f"No SVG templates found for school: {school_name_raw}"
        )

    first_name = _normalise_string(record.get("first_name"))
    last_name = _normalise_string(record.get("last_name"))
    child_output_base = _build_child_output_base(first_name, last_name, school_name_raw)

    school_output_dir = output_root / school_id
    child_output_dir = school_output_dir / child_output_base

    _ensure_directory(child_output_dir)

    working_dir = _get_temp_working_dir(outer_prefix, ID_CARD_WORKING_SUBDIR)
    _prepare_working_directory(template_dir, working_dir)
    photos_output_dir = working_dir / "images"

    class_name = _normalise_string(record.get("class_name")).upper()
    blood_group = _normalise_string(record.get("blood_group"))
    age = _normalise_string(record.get("age"))
    date_of_birth = _normalise_string(record.get("date_of_birth"))
    if date_of_birth:
        date_of_birth = _format_date(date_of_birth)

    gender = _normalise_string(record.get("gender"))
    current_address = _normalise_string(record.get("current_address"))

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
        front_svg_path = _working_file_path(front_template, template_dir, working_dir)
        _process_svg(front_svg_path, text_updates_front, image_updates_front, address_updates_front)
        front_pdf_path = child_output_dir / f"{child_output_base}_FRONT.pdf"
        callInkscape(str(front_svg_path), str(front_pdf_path))
        generated = True

    if back_template is not None:
        back_svg_path = _working_file_path(back_template, template_dir, working_dir)
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
