from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

import msg_to_pdf_dropzone.pdf_writer as pdf_writer
from msg_to_pdf_dropzone.models import EmailRecord
from msg_to_pdf_dropzone.pdf_writer import build_email_html_document, write_email_pdf
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
        html_body="",
        attachment_names=["doc.txt"],
        thread_key=normalize_thread_subject("Test Subject"),
    )

    output_file = tmp_path / "test.pdf"
    write_email_pdf(record, output_file)

    assert output_file.exists()
    assert output_file.stat().st_size > 0


def test_build_email_html_document_prefers_html_body_fragment() -> None:
    record = EmailRecord(
        source_path=Path("sample.msg"),
        subject="Sample Subject",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Plain text fallback",
        html_body="<html><body><p><b>Rich</b> body</p></body></html>",
        attachment_names=["file-a.xlsx"],
        thread_key=normalize_thread_subject("Sample Subject"),
    )

    html_document = build_email_html_document(record)

    assert "<b>Rich</b> body" in html_document
    assert "Sample Subject" in html_document
    assert "file-a.xlsx" in html_document


def test_write_email_pdf_prefers_outlook_edge_pipeline(monkeypatch, tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=tmp_path / "sample.msg",
        subject="Sample Subject",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Plain text fallback",
        html_body="<html><body><p><b>Rich</b> body</p></body></html>",
        attachment_names=[],
        thread_key=normalize_thread_subject("Sample Subject"),
    )
    record.source_path.write_text("dummy", encoding="utf-8")

    def fake_office_pipeline(_: Path, output_path: Path) -> bool:
        output_path.write_bytes(b"%PDF-1.4\nfake\n")
        return True

    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_outlook_and_edge", fake_office_pipeline)
    monkeypatch.setattr(
        pdf_writer,
        "_try_write_pdf_via_edge_html",
        lambda _html, _output: (_ for _ in ()).throw(AssertionError("Edge HTML pipeline should not run.")),
    )

    output_file = tmp_path / "test.pdf"
    write_email_pdf(record, output_file)

    assert output_file.exists()
    assert output_file.stat().st_size > 0
