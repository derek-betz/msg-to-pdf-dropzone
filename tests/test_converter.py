from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from msg_to_pdf_dropzone.converter import ConversionError, convert_msg_files
from msg_to_pdf_dropzone.models import EmailRecord
from msg_to_pdf_dropzone.thread_logic import normalize_thread_subject


def test_convert_msg_files_enforces_batch_limit(tmp_path: Path) -> None:
    too_many = [tmp_path / f"{index}.msg" for index in range(11)]
    try:
        convert_msg_files(too_many, tmp_path)
    except ConversionError as exc:
        assert "up to 10 files" in str(exc)
    else:
        raise AssertionError("Expected ConversionError for >10 files")


def test_convert_msg_files_uses_latest_thread_date_for_filename(monkeypatch, tmp_path: Path) -> None:
    email_one = tmp_path / "one.msg"
    email_two = tmp_path / "two.msg"
    email_one.touch()
    email_two.touch()

    older = EmailRecord(
        source_path=email_one,
        subject="Re: Sample Thread",
        sent_at=datetime(2026, 1, 1, tzinfo=timezone.utc),
        sender="a@example.com",
        to="b@example.com",
        cc="",
        body="older",
        attachment_names=[],
        thread_key=normalize_thread_subject("Re: Sample Thread"),
    )
    newer = EmailRecord(
        source_path=email_two,
        subject="Sample Thread",
        sent_at=datetime(2026, 1, 5, tzinfo=timezone.utc),
        sender="a@example.com",
        to="b@example.com",
        cc="",
        body="newer",
        attachment_names=[],
        thread_key=normalize_thread_subject("Sample Thread"),
    )

    records = {email_one: older, email_two: newer}

    def fake_parse(path: Path) -> EmailRecord:
        return records[path]

    def fake_write(record: EmailRecord, output_path: Path) -> Path:
        output_path.write_text(record.body, encoding="utf-8")
        return output_path

    monkeypatch.setattr("msg_to_pdf_dropzone.converter.parse_msg_file", fake_parse)
    monkeypatch.setattr("msg_to_pdf_dropzone.converter.write_email_pdf", fake_write)

    result = convert_msg_files([email_one, email_two], tmp_path)

    assert len(result.converted_files) == 2
    assert all(path.name.startswith("2026-01-05 ") for path in result.converted_files)
