from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from msg_to_pdf_dropzone.models import EmailRecord
from msg_to_pdf_dropzone.pdf_writer import write_email_pdf
from msg_to_pdf_dropzone.thread_logic import normalize_thread_subject


def test_write_email_pdf_creates_nonempty_file(tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=Path("sample.msg"),
        subject="Test Subject",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="cc@example.com",
        body="Hello\n\nThis is a test body.",
        attachment_names=["doc.txt"],
        thread_key=normalize_thread_subject("Test Subject"),
    )

    output_file = tmp_path / "test.pdf"
    write_email_pdf(record, output_file)

    assert output_file.exists()
    assert output_file.stat().st_size > 0
