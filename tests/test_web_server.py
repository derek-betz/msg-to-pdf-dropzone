from __future__ import annotations

import asyncio
from io import BytesIO
import json
from pathlib import Path

from fastapi.testclient import TestClient

from msg_to_pdf_dropzone.converter import ConversionResult
from msg_to_pdf_dropzone.task_events import emit_task_event
from msg_to_pdf_dropzone.web_server import EventBroker, create_app, publish_preview_sequence


def test_health_and_queue_snapshot(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    client = TestClient(create_app())

    health = client.get("/api/health")
    queue = client.get("/api/queue")

    assert health.status_code == 200
    assert health.json()["ok"] is True
    assert queue.status_code == 200
    assert queue.json()["items"] == []


def test_upload_and_convert_flow(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    converted_path = tmp_path / "output" / "2026-04-04_Sample.pdf"

    def fake_convert(msg_paths, output_dir, event_sink=None, task_ids_by_source_path=None, **_kwargs) -> ConversionResult:
        assert len(msg_paths) == 1
        assert msg_paths[0].suffix.lower() == ".msg"
        assert output_dir == tmp_path / "output"
        converted_path.parent.mkdir(parents=True, exist_ok=True)
        converted_path.write_text("pdf", encoding="utf-8")
        task_id = task_ids_by_source_path[msg_paths[0].resolve()]
        emit_task_event(
            event_sink,
            task_id=task_id,
            stage="complete",
            file_name=msg_paths[0].name,
            success=True,
            meta={"outputPath": str(converted_path)},
        )
        return ConversionResult(
            requested_count=1,
            converted_files=[converted_path],
            skipped_files=[],
            errors=[],
            total_seconds=0.25,
            timing_lines=["Total 0.25s"],
        )

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.convert_msg_files", fake_convert)
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[("files", ("sample.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
    )
    assert upload.status_code == 200
    queued_item = upload.json()["items"][0]
    assert queued_item["taskId"].startswith("msg-to-pdf-")

    convert = client.post(
        "/api/convert",
        json={"ids": [queued_item["id"]], "output_dir": str(tmp_path / "output")},
    )
    assert convert.status_code == 200
    assert convert.json()["convertedFiles"] == [str(converted_path)]

    queue = client.get("/api/queue")
    assert queue.status_code == 200
    items = queue.json()["items"]
    assert len(items) == 1
    assert items[0]["id"] == queued_item["id"]
    assert items[0]["stage"] == "complete"


def test_upload_accepts_outlook_source_hint(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[("files", ("sample.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
        data={"source_hint": "outlook"},
    )
    assert upload.status_code == 200
    payload = upload.json()
    assert len(payload["accepted"]) == 1
    assert payload["accepted"][0]["source"] == "outlook"
    assert payload["items"][0]["source"] == "outlook"


def test_completed_items_stay_visible_without_blocking_new_uploads(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")

    def fake_convert(msg_paths, output_dir, event_sink=None, task_ids_by_source_path=None, **_kwargs) -> ConversionResult:
        converted_files = []
        for msg_path in msg_paths:
            output_path = output_dir / f"{Path(msg_path).stem}.pdf"
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text("pdf", encoding="utf-8")
            converted_files.append(output_path)
            task_id = task_ids_by_source_path[msg_path.resolve()]
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="complete",
                file_name=msg_path.name,
                success=True,
                meta={"outputPath": str(output_path)},
            )
        return ConversionResult(
            requested_count=len(msg_paths),
            converted_files=converted_files,
            skipped_files=[],
            errors=[],
            total_seconds=0.5,
            timing_lines=["Total 0.50s"],
        )

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.convert_msg_files", fake_convert)
    client = TestClient(create_app())

    first_upload = client.post(
        "/api/upload",
        files=[("files", ("first.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
    )
    assert first_upload.status_code == 200
    first_item = first_upload.json()["items"][0]

    convert = client.post(
        "/api/convert",
        json={"ids": [first_item["id"]], "output_dir": str(tmp_path / "output")},
    )
    assert convert.status_code == 200

    second_upload = client.post(
        "/api/upload",
        files=[("files", ("second.msg", BytesIO(b"msg-bytes-2"), "application/vnd.ms-outlook"))],
    )
    assert second_upload.status_code == 200
    assert len(second_upload.json()["accepted"]) == 1


def test_upload_recreates_missing_staging_dir(monkeypatch, tmp_path: Path) -> None:
    staging_dir = tmp_path / "staging"
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", staging_dir)
    client = TestClient(create_app())

    if staging_dir.exists():
        staging_dir.rmdir()

    upload = client.post(
        "/api/upload",
        files=[("files", ("sample.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
    )

    assert upload.status_code == 200
    payload = upload.json()
    assert payload["rejectedCount"] == 0
    assert len(payload["accepted"]) == 1
    assert staging_dir.exists()


def test_publish_preview_sequence_emits_stage_events(monkeypatch) -> None:
    async def fake_sleep(_seconds: float) -> None:
        return None

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.asyncio.sleep", fake_sleep)
    broker = EventBroker()
    subscriber = broker.subscribe()

    asyncio.run(publish_preview_sequence(broker, pipeline="edge_html"))

    payloads = []
    while not subscriber.empty():
        payloads.append(json.loads(subscriber.get_nowait()))

    stages = [payload["stage"] for payload in payloads if "stage" in payload]
    assert stages[0] == "drop_received"
    assert stages[-1] == "complete"
    assert any(payload.get("pipeline") == "edge_html" for payload in payloads if "stage" in payload)
