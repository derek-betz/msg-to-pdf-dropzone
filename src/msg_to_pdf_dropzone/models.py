from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(slots=True)
class EmailRecord:
    source_path: Path
    subject: str
    sent_at: datetime
    sender: str
    to: str
    cc: str
    body: str
    attachment_names: list[str]
    thread_key: str
