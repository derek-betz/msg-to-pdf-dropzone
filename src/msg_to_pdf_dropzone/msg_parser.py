from __future__ import annotations

from datetime import datetime
from email.utils import parsedate_to_datetime
import mimetypes
from pathlib import Path

import extract_msg

from .models import EmailRecord, InlineImageAsset
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


def _as_html(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        for encoding in ("utf-8", "utf-16", "latin-1"):
            try:
                return value.decode(encoding).strip()
            except UnicodeDecodeError:
                continue
        return value.decode("utf-8", errors="ignore").strip()
    return str(value).strip()


def _as_bytes(value: object) -> bytes:
    if value is None:
        return b""
    if isinstance(value, bytes):
        return value
    if isinstance(value, bytearray):
        return bytes(value)
    if isinstance(value, memoryview):
        return value.tobytes()
    try:
        return bytes(value)
    except Exception:
        return b""


def _normalize_cid(value: str) -> str:
    cid = (value or "").strip().strip("<>").strip()
    if cid.lower().startswith("cid:"):
        cid = cid[4:].strip()
    return cid


def _guess_image_mime(filename: str) -> str:
    if not filename:
        return ""
    guessed, _ = mimetypes.guess_type(filename, strict=False)
    if guessed and guessed.lower().startswith("image/"):
        return guessed
    return ""


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
        inline_images: list[InlineImageAsset] = []
        seen_cids: set[str] = set()
        for attachment in getattr(message, "attachments", []):
            name = (
                _as_text(getattr(attachment, "longFilename", ""))
                or _as_text(getattr(attachment, "filename", ""))
                or "unnamed-attachment"
            )
            attachment_names.append(name)

            cid = _normalize_cid(
                _as_text(getattr(attachment, "cid", ""))
                or _as_text(getattr(attachment, "contentId", ""))
            )
            if not cid:
                continue

            cid_key = cid.lower()
            if cid_key in seen_cids:
                continue

            mime_type = (
                _as_text(getattr(attachment, "mimetype", ""))
                or _as_text(getattr(attachment, "mimeType", ""))
                or _guess_image_mime(name)
            ).lower()
            if not mime_type.startswith("image/"):
                continue

            data = _as_bytes(getattr(attachment, "data", None))
            if not data:
                continue

            inline_images.append(
                InlineImageAsset(
                    cid=cid,
                    mime_type=mime_type,
                    filename=name,
                    data=data,
                    size_bytes=len(data),
                )
            )
            seen_cids.add(cid_key)

        body = _as_text(getattr(message, "body", ""))
        if not body:
            body = "(No plain text body found in source email.)"
        html_body = (
            _as_html(getattr(message, "htmlBody", None))
            or _as_html(getattr(message, "html_body", None))
            or _as_html(getattr(message, "html", None))
        )

        return EmailRecord(
            source_path=msg_path,
            subject=subject,
            sent_at=sent_at,
            sender=_as_text(getattr(message, "sender", "")),
            to=_as_text(getattr(message, "to", "")),
            cc=_as_text(getattr(message, "cc", "")),
            body=body,
            html_body=html_body,
            attachment_names=attachment_names,
            thread_key=normalize_thread_subject(subject),
            inline_images=inline_images,
        )
    finally:
        message.close()
