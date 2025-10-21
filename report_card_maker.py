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
    _build_school_folder_name,
    _candidate_template_directories,
    _copy_photo,
    _ensure_directory,
    _extract_outer_code_prefix,
    _format_date,
    _get_temp_working_dir,
    _guardian_type,
    _is_missing,
    _normalise_string,
    _prepare_working_directory,
    _process_svg,
    _sanitize_filename_component,
    _working_file_path,
    clean_branch_name,
    custom_title_case,
)

DEFAULT_TEMPLATE_ROOT = Path(r"\\pixartnas\home\INTERNAL_PROCESSING\ALL REPORT CARD SRC")
DEFAULT_OUTPUT_ROOT = Path("Report cards")
REPORT_CARD_WORKING_SUBDIR = "ReportCardWorking"


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
        return class_specific
    for extension in (".svg", ".SVG"):
        candidate = template_dir / f"{prefix}_{class_name}{extension}"
        if candidate.exists():
            return candidate


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

    outer_prefix = _extract_outer_code_prefix(record.get("outer_code"))
    if not outer_prefix:
        return False

    template_group_dir = template_root / outer_prefix
    if not template_group_dir.is_dir():
        raise TemplateNotFoundError(
            f"Template directory not found for outer code prefix: {outer_prefix}"
        )

    class_name = _normalise_string(record.get("class_name"))
    class_name_for_front = class_name.replace("FRONT_", "")

    template_dir: Optional[Path] = None
    front_template: Optional[Path] = None
    back_template: Optional[Path] = None

    for candidate in _candidate_template_directories(
        template_group_dir, school_id, school_name_raw
    ):
        front_candidate = _resolve_template(candidate, "FRONT", class_name_for_front)
        back_candidate = _resolve_template(candidate, "BACK", class_name)
        if front_candidate is None and back_candidate is None:
            continue
        template_dir = candidate
        front_template = front_candidate
        back_template = back_candidate
        break

    if template_dir is None:
        raise TemplateNotFoundError(f"No SVG templates found for school: {school_name_raw}")

    first_name = _normalise_string(record.get("first_name"))
    last_name = _normalise_string(record.get("last_name"))
    child_output_base = _build_child_output_base(first_name, last_name, school_name_raw)

    school_folder_name = _build_school_folder_name(school_id, school_name_raw)
    school_output_dir = output_root / school_folder_name
    child_output_dir = school_output_dir / child_output_base

    _ensure_directory(child_output_dir)

    working_dir = _get_temp_working_dir(outer_prefix, REPORT_CARD_WORKING_SUBDIR)
    _prepare_working_directory(template_dir, working_dir)
    photos_output_dir = working_dir / "images"

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
