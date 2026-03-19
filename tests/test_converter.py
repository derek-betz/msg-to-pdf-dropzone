from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from msg_to_pdf_dropzone.converter import ConversionError, convert_msg_files
from msg_to_pdf_dropzone.models import EmailRecord
from msg_to_pdf_dropzone.task_events import emit_task_event
from msg_to_pdf_dropzone.thread_logic import get_latest_thread_dates, normalize_thread_subject


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
        html_body="",
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
        html_body="",
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
    expected_latest = get_latest_thread_dates([older, newer])[newer.thread_key]
    expected_prefix = f"{expected_latest:%Y-%m-%d}_"
    assert all(path.name.startswith(expected_prefix) for path in result.converted_files)


def test_convert_msg_files_populates_timing_lines(monkeypatch, tmp_path: Path) -> None:
    email_path = tmp_path / "one.msg"
    email_path.touch()

    record = EmailRecord(
        source_path=email_path,
        subject="Timing Test",
        sent_at=datetime(2026, 1, 6, tzinfo=timezone.utc),
        sender="a@example.com",
        to="b@example.com",
        cc="",
        body="body",
        html_body="",
        attachment_names=[],
        thread_key=normalize_thread_subject("Timing Test"),
    )

    def fake_parse(path: Path) -> EmailRecord:
        return record

    def fake_write(_record: EmailRecord, output_path: Path, *, diagnostics=None) -> Path:
        output_path.write_text("ok", encoding="utf-8")
        if diagnostics is not None:
            diagnostics.pipeline = "fake_pipeline"
            diagnostics.stage_seconds["fake_stage"] = 0.01
            diagnostics.image_metrics["total_images"] = 2
            diagnostics.image_metrics["cid_resolved"] = 1
            diagnostics.total_seconds = 0.01
        return output_path

    monkeypatch.setattr("msg_to_pdf_dropzone.converter.parse_msg_file", fake_parse)
    monkeypatch.setattr("msg_to_pdf_dropzone.converter.write_email_pdf", fake_write)

    result = convert_msg_files([email_path], tmp_path)

    assert result.total_seconds >= 0.0
    assert result.parse_seconds >= 0.0
    assert result.write_seconds >= 0.0
    assert result.timing_lines
    assert "Total " in result.timing_lines[0]
    assert any("one.msg:" in line for line in result.timing_lines)
    assert any("pipeline fake_pipeline" in line for line in result.timing_lines)
    assert any("images total_images 2" in line for line in result.timing_lines)
    assert len(result.file_timing_records) == 1
    assert result.file_timing_records[0].file_name == "one.msg"
    assert result.file_timing_records[0].pipeline == "fake_pipeline"
    assert result.file_timing_records[0].stage_seconds["fake_stage"] == 0.01
    assert result.file_timing_records[0].image_metrics["total_images"] == 2


def test_convert_msg_files_emits_task_events(monkeypatch, tmp_path: Path) -> None:
    email_path = tmp_path / "one.msg"
    email_path.touch()

    record = EmailRecord(
        source_path=email_path,
        subject="Event Test",
        sent_at=datetime(2026, 1, 6, tzinfo=timezone.utc),
        sender="a@example.com",
        to="b@example.com",
        cc="",
        body="body",
        html_body="",
        attachment_names=[],
        thread_key=normalize_thread_subject("Event Test"),
    )

    def fake_parse(path: Path) -> EmailRecord:
        return record

    def fake_write(
        _record: EmailRecord,
        output_path: Path,
        *,
        diagnostics=None,
        event_sink=None,
        task_id=None,
    ) -> Path:
        output_path.write_text("ok", encoding="utf-8")
        if diagnostics is not None:
            diagnostics.pipeline = "edge_html"
        emit_task_event(
            event_sink,
            task_id=task_id or "fallback-task-id",
            stage="pipeline_selected",
            file_name=_record.source_path.name,
            pipeline="edge_html",
        )
        return output_path

    monkeypatch.setattr("msg_to_pdf_dropzone.converter.parse_msg_file", fake_parse)
    monkeypatch.setattr("msg_to_pdf_dropzone.converter.write_email_pdf", fake_write)

    events = []
    result = convert_msg_files(
        [email_path],
        tmp_path,
        event_sink=events.append,
        task_ids_by_source_path={email_path.resolve(): "task-123"},
    )

    assert result.converted_files
    assert [event.stage for event in events] == [
        "parse_started",
        "filename_built",
        "pdf_pipeline_started",
        "pipeline_selected",
        "pdf_written",
        "deliver_started",
        "complete",
    ]
    assert events[0].task_id == "task-123"
    assert events[3].pipeline == "edge_html"
    assert events[-1].success is True


def test_convert_msg_files_preserves_batch_meta_across_events(monkeypatch, tmp_path: Path) -> None:
    email_path = tmp_path / "bundle.msg"
    email_path.touch()

    record = EmailRecord(
        source_path=email_path,
        subject="Bundle Test",
        sent_at=datetime(2026, 1, 6, tzinfo=timezone.utc),
        sender="a@example.com",
        to="b@example.com",
        cc="",
        body="body",
        html_body="",
        attachment_names=[],
        thread_key=normalize_thread_subject("Bundle Test"),
    )

    def fake_parse(path: Path) -> EmailRecord:
        return record

    def fake_write(
        _record: EmailRecord,
        output_path: Path,
        *,
        diagnostics=None,
        event_sink=None,
        task_id=None,
        event_meta=None,
    ) -> Path:
        output_path.write_text("ok", encoding="utf-8")
        if diagnostics is not None:
            diagnostics.pipeline = "edge_html"
        emit_task_event(
            event_sink,
            task_id=task_id or "fallback-task-id",
            stage="pipeline_selected",
            file_name=_record.source_path.name,
            pipeline="edge_html",
            meta=event_meta,
        )
        return output_path

    monkeypatch.setattr("msg_to_pdf_dropzone.converter.parse_msg_file", fake_parse)
    monkeypatch.setattr("msg_to_pdf_dropzone.converter.write_email_pdf", fake_write)

    events = []
    batch_meta = {
        email_path.resolve(): {
            "batchId": "msg-batch-001",
            "batchSize": 3,
            "batchIndex": 2,
        }
    }
    convert_msg_files(
        [email_path],
        tmp_path,
        event_sink=events.append,
        task_ids_by_source_path={email_path.resolve(): "task-bundle"},
        batch_meta_by_source_path=batch_meta,
    )

    assert events
    assert all(event.meta is not None for event in events)
    assert all(event.meta["batchId"] == "msg-batch-001" for event in events if event.meta is not None)
    assert all(event.meta["batchSize"] == 3 for event in events if event.meta is not None)
