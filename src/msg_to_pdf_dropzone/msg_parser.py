from __future__ import annotations

from datetime import datetime
from email.utils import parsedate_to_datetime
from pathlib import Path

import extract_msg

from .models import EmailRecord
from .thread_logic import normalize_thread_subject


def _parse_sent_date(raw_date: object) -> datetime | None:
    if raw_date is None:
        return None
    if isinstance(raw_date, datetime):
        return raw_date
    raw_value = str(raw_date).strip()
    if not raw_value:
        return None
    try:
        return parsedate_to_datetime(raw_value)
    except (TypeError, ValueError):
        return None


def _as_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_msg_file(msg_path: Path) -> EmailRecord:
    if msg_path.suffix and msg_path.suffix.lower() != ".msg":
        raise ValueError(f"Unsupported file type: {msg_path}")

    message = extract_msg.Message(str(msg_path))
    try:
        subject = _as_text(getattr(message, "subject", "")) or "No Subject"
        sent_at = _parse_sent_date(getattr(message, "date", None))
        if sent_at is None:
            sent_at = datetime.fromtimestamp(msg_path.stat().st_mtime).astimezone()
        elif sent_at.tzinfo is None:
            sent_at = sent_at.astimezone()

        attachment_names: list[str] = []
        for attachment in getattr(message, "attachments", []):
            name = (
                _as_text(getattr(attachment, "longFilename", ""))
                or _as_text(getattr(attachment, "filename", ""))
                or "unnamed-attachment"
            )
            attachment_names.append(name)

        body = _as_text(getattr(message, "body", ""))
        if not body:
            body = "(No plain text body found in source email.)"

        return EmailRecord(
            source_path=msg_path,
            subject=subject,
            sent_at=sent_at,
            sender=_as_text(getattr(message, "sender", "")),
            to=_as_text(getattr(message, "to", "")),
            cc=_as_text(getattr(message, "cc", "")),
            body=body,
            attachment_names=attachment_names,
            thread_key=normalize_thread_subject(subject),
        )
    finally:
        message.close()
