"""Utilities for generating personalised report card PDFs from SVG templates."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, Iterator, Optional, Sequence, Set, Tuple

from doc_maker import callInkscape
from id_card_maker import (
    DEFAULT_PHOTO_ROOT,
    TemplateNotFoundError,
    _build_child_output_base,
    _copy_photo,
    _ensure_directory,
    _format_date,
    _guardian_type,
    _is_missing,
    _normalise_string,
    _prepare_working_directory,
    _process_svg,
    _sanitize_filename_component,
    clean_branch_name,
    custom_title_case,
)

print(Path(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL REPORT CARD SRC\149\FRONT_UKG.svg").exists())

DEFAULT_TEMPLATE_ROOT = Path(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL REPORT CARD SRC")
DEFAULT_OUTPUT_ROOT = Path("Report cards")


def _generate_class_variants(class_name: str) -> Sequence[str]:
    normalized = _normalise_string(class_name)
    if not normalized:
        return []

    sanitized = _sanitize_filename_component(normalized, normalized)
    replacements = [
        normalized,
        normalized.replace(" ", "_"),
        normalized.replace(" ", ""),
        normalized.replace("-", "_"),
        normalized.replace("-", ""),
        normalized.replace("/", "_"),
        normalized.replace("/", ""),
        sanitized,
        sanitized.replace("_", ""),
    ]

    variants = []
    seen = set()
    for candidate in replacements:
        candidate = re.sub(r"_+", "_", candidate).strip("_")
        if not candidate:
            continue
        for variant in (candidate, candidate.upper(), candidate.lower()):
            if variant not in seen:
                variants.append(variant)
                seen.add(variant)
    return variants


def _find_class_template(template_dir: Path, prefix: str, class_name: str) -> Optional[Path]:
    for variant in _generate_class_variants(class_name):
        candidate = template_dir / f"{prefix}_{variant}.svg"
        if candidate.exists():
            return candidate
        candidate_upper = template_dir / f"{prefix}_{variant}.SVG"
        if candidate_upper.exists():
            return candidate_upper
    return None


def _resolve_template(template_dir: Path, prefix: str, class_name: str) -> Optional[Path]:
    class_specific = _find_class_template(template_dir, prefix, class_name)
    if class_specific is not None:
        print("class_specific", class_specific)
        return class_specific
    for extension in (".svg", ".SVG"):
        candidate = template_dir / f"{prefix}_{class_name}{extension}"
        print("candidate", candidate)
        if candidate.exists():
            return candidate


def _normalise_directory_name(value: str) -> str:
    return re.sub(r"[^0-9a-z]+", "_", value.lower()).strip("_")


def _resolve_template_directory(
    template_root: Path, school_id: str, school_name: str
) -> Optional[Path]:
    """Locate the template directory for a school.

    Some schools store their templates in folders that include either the
    school name, the school ID, or both.  When the direct ``school_id``
    lookup fails we attempt to locate a directory whose normalised name
    contains the required pieces of information.
    """

    normalised_school_id = _normalise_string(school_id)
    normalised_school_name = _normalise_string(school_name)
    print("template", DEFAULT_TEMPLATE_ROOT)
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


def personalize_report_card(
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
    print("template dir", template_dir)
    if template_dir is None:
        raise TemplateNotFoundError(
            f"Template directory not found for school: {school_id}"
        )
    

    class_name = _normalise_string(record.get("class_name"))
    front_template = _resolve_template(template_dir, "FRONT", class_name.replace("FRONT_", ""))
    back_template = _resolve_template(template_dir, "BACK", class_name)

    print("front template", front_template)
    print("back template", back_template)

    if not front_template and not back_template:
        raise TemplateNotFoundError(f"No SVG templates found for school: {school_name_raw}")

    first_name = _normalise_string(record.get("first_name"))
    last_name = _normalise_string(record.get("last_name"))
    child_output_base = _build_child_output_base(first_name, last_name, school_name_raw)

    school_folder_name = f"{_sanitize_filename_component(school_id, 'school')}_{_sanitize_filename_component(school_name_raw, 'school')}"
    school_output_dir = output_root / school_folder_name
    child_output_dir = school_output_dir / child_output_base
    photos_output_dir = child_output_dir / "working" / "images"

    _ensure_directory(child_output_dir)
    working_dir = child_output_dir / "working"
    _prepare_working_directory(template_dir, working_dir)

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
        "grade": (class_name.upper(), 14, 0.5),
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


def generate_report_cards(
    records: Iterable[Dict[str, object]],
    *,
    template_root: Path = DEFAULT_TEMPLATE_ROOT,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
    photo_root: Path = DEFAULT_PHOTO_ROOT,
) -> int:
    count = 0
    seen_children: Set[Tuple[str, str]] = set()
    for record in records:
        dedupe_key = (
            _normalise_string(record.get("school_id")),
            _normalise_string(record.get("user_id")),
        )
        if dedupe_key[1]:
            if dedupe_key in seen_children:
                continue
            seen_children.add(dedupe_key)
        try:
            if personalize_report_card(
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

print("template", DEFAULT_TEMPLATE_ROOT)


def load_records_from_workbook(workbook_path: Path) -> Iterator[Dict[str, object]]:
    import pandas as pd

    df = pd.read_excel(workbook_path, header=0)
    for record in df.to_dict("records"):
        yield record


def generate_report_cards_from_workbook(
    workbook_path: Path,
    *,
    template_root: Path = DEFAULT_TEMPLATE_ROOT,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
    photo_root: Path = DEFAULT_PHOTO_ROOT,
) -> int:
    return generate_report_cards(
        load_records_from_workbook(workbook_path),
        template_root=template_root,
        output_root=output_root,
        photo_root=photo_root,
    )


def _parse_args(argv: Optional[Sequence[str]] = None) -> Tuple[Path, Path, Path, Path]:
    import argparse

    parser = argparse.ArgumentParser(description="Generate report cards from an Excel workbook.")
    parser.add_argument("workbook_path", type=Path, help="Path to the Excel workbook containing report card data")
    parser.add_argument(
        "--template-root",
        type=Path,
        default=DEFAULT_TEMPLATE_ROOT,
        help="Directory containing per-school report card SVG templates",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=DEFAULT_OUTPUT_ROOT,
        help="Directory where generated report cards will be written",
    )
    parser.add_argument(
        "--photo-root",
        type=Path,
        default=DEFAULT_PHOTO_ROOT,
        help="Directory containing student and guardian photos",
    )
    args = parser.parse_args(argv)
    return args.workbook_path, args.template_root, args.output_root, args.photo_root


def main(argv: Optional[Sequence[str]] = None) -> int:
    workbook_path, template_root, output_root, photo_root = _parse_args(argv)
    count = generate_report_cards_from_workbook(
        workbook_path,
        template_root=template_root,
        output_root=output_root,
        photo_root=photo_root,
    )
    print(f"Generated {count} report card(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
