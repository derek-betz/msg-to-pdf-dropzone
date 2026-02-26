from __future__ import annotations

import re
from datetime import date
from pathlib import Path

from .models import EmailRecord

INVALID_FILENAME_CHARS = re.compile(r"[<>:\"/\\|?*\x00-\x1F]")
WHITESPACE_PATTERN = re.compile(r"\s+")
THREAD_PREFIX_PATTERN = re.compile(r"^\s*((re|fw|fwd)\s*:\s*)+", re.IGNORECASE)


def normalize_thread_subject(subject: str) -> str:
    clean_subject = (subject or "").strip()
    clean_subject = THREAD_PREFIX_PATTERN.sub("", clean_subject).strip().lower()
    return clean_subject or "no-subject"


def get_latest_thread_dates(records: list[EmailRecord]) -> dict[str, date]:
    latest_dates: dict[str, date] = {}
    for record in records:
        message_date = record.sent_at.date()
        current = latest_dates.get(record.thread_key)
        if current is None or message_date > current:
            latest_dates[record.thread_key] = message_date
    return latest_dates


def sanitize_filename_part(value: str, max_length: int = 120) -> str:
    candidate = (value or "").strip() or "No Subject"
    candidate = INVALID_FILENAME_CHARS.sub("_", candidate)
    candidate = WHITESPACE_PATTERN.sub(" ", candidate).strip(" .")
    if not candidate:
        candidate = "No Subject"
    return candidate[:max_length].rstrip(" .")


def build_pdf_filename(subject: str, thread_latest_date: date) -> str:
    safe_subject = sanitize_filename_part(subject)
    return f"{thread_latest_date:%Y-%m-%d} {safe_subject}.pdf"


def make_unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 2
    stem = path.stem
    suffix = path.suffix
    while True:
        candidate = path.with_name(f"{stem} ({counter}){suffix}")
        if not candidate.exists():
            return candidate
        counter += 1
