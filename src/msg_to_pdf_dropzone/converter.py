from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .msg_parser import parse_msg_file
from .pdf_writer import write_email_pdf
from .thread_logic import build_pdf_filename, get_latest_thread_dates, make_unique_path

MAX_FILES_PER_BATCH = 10


class ConversionError(Exception):
    pass


@dataclass(slots=True)
class ConversionResult:
    requested_count: int = 0
    converted_files: list[Path] = field(default_factory=list)
    skipped_files: list[Path] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def convert_msg_files(msg_paths: list[Path], output_dir: Path) -> ConversionResult:
    result = ConversionResult(requested_count=len(msg_paths))
    if not msg_paths:
        raise ConversionError("No .msg files were provided.")
    if len(msg_paths) > MAX_FILES_PER_BATCH:
        raise ConversionError(f"Select up to {MAX_FILES_PER_BATCH} files per batch.")

    email_records = []
    for msg_path in msg_paths:
        if msg_path.suffix.lower() != ".msg":
            result.skipped_files.append(msg_path)
            result.errors.append(f"Skipped non-msg file: {msg_path.name}")
            continue
        try:
            email_records.append(parse_msg_file(msg_path))
        except Exception as exc:  # pragma: no cover - external parser failure branch
            result.skipped_files.append(msg_path)
            result.errors.append(f"Failed to parse {msg_path.name}: {exc}")

    if not email_records:
        return result

    latest_thread_dates = get_latest_thread_dates(email_records)
    for record in email_records:
        try:
            file_name = build_pdf_filename(record.subject, latest_thread_dates[record.thread_key])
            output_path = make_unique_path(output_dir / file_name)
            write_email_pdf(record, output_path)
            result.converted_files.append(output_path)
        except Exception as exc:  # pragma: no cover - external PDF writer failure branch
            result.skipped_files.append(record.source_path)
            result.errors.append(f"Failed to convert {record.source_path.name}: {exc}")

    return result
