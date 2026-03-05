from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


@dataclass(slots=True)
class InlineImageAsset:
    cid: str
    mime_type: str
    filename: str
    data: bytes
    size_bytes: int


@dataclass(slots=True)
class EmailRecord:
    source_path: Path
    subject: str
    sent_at: datetime
    sender: str
    to: str
    cc: str
    body: str
    html_body: str
    attachment_names: list[str]
    thread_key: str
    inline_images: list[InlineImageAsset] = field(default_factory=list)
