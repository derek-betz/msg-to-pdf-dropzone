from __future__ import annotations

import asyncio
from datetime import datetime, timezone
from io import BytesIO
import json
from pathlib import Path
import re

from fastapi.testclient import TestClient

import msg_to_pdf_dropzone.web_server as web_server
from msg_to_pdf_dropzone.converter import ConversionResult
from msg_to_pdf_dropzone.models import EmailRecord
from msg_to_pdf_dropzone.task_events import emit_task_event
from msg_to_pdf_dropzone.thread_logic import normalize_thread_subject
from msg_to_pdf_dropzone.web_server import EventBroker, create_app, publish_preview_sequence


WEB_UI_DIR = Path(__file__).resolve().parents[1] / "src" / "msg_to_pdf_dropzone" / "web_ui"
WEB_UI_APP_PATH = WEB_UI_DIR / "app.js"
WEB_UI_CSS_PATH = WEB_UI_DIR / "app.css"
WEB_UI_DROPZONE_CONTROLLER_PATH = WEB_UI_DIR / "dropzone_controller.js"


def test_health_and_queue_snapshot(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    client = TestClient(create_app())

    health = client.get("/api/health")
    queue = client.get("/api/queue")

    assert health.status_code == 200
    assert health.json()["ok"] is True
    assert queue.status_code == 200
    assert queue.json()["items"] == []


def test_health_and_version_include_release_metadata(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    monkeypatch.setenv("MSG_TO_PDF_APP_REVISION", "abc123")
    client = TestClient(create_app())

    health = client.get("/api/health")
    version = client.get("/api/version")

    assert health.status_code == 200
    assert version.status_code == 200
    assert health.json()["appName"] == "msg-to-pdf-dropzone"
    assert health.json()["sourceRevision"] == "abc123"
    assert version.json()["sourceRevision"] == "abc123"
    assert isinstance(version.json()["appVersion"], str)


def test_version_payload_reads_deploy_release_file(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.delenv("MSG_TO_PDF_APP_REVISION", raising=False)
    monkeypatch.delenv("MSG_TO_PDF_SOURCE_REVISION", raising=False)
    monkeypatch.setattr(web_server, "PACKAGE_ROOT", tmp_path)
    (tmp_path / "_release.json").write_bytes(
        "\ufeff".encode("utf-8")
        + json.dumps(
            {
                "sourceRevision": "deployed-revision",
                "sourceBranch": "main",
                "deployedAt": "2026-05-08T00:00:00Z",
                "deployedBy": "HANSON\\betz02340",
            }
        ).encode("utf-8")
    )

    payload = web_server.build_version_payload()

    assert payload["sourceRevision"] == "deployed-revision"
    assert payload["sourceBranch"] == "main"
    assert payload["deployedAt"] == "2026-05-08T00:00:00Z"
    assert payload["deployedBy"] == "HANSON\\betz02340"


def test_index_includes_feedback_modal(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    client = TestClient(create_app())

    response = client.get("/")

    assert response.status_code == 200
    assert "Send Feedback" in response.text
    assert "Add MSG files" in response.text
    assert "Confirm names" in response.text
    assert "Convert and download" in response.text
    assert "PDFs download individually through the browser." in response.text
    assert "Choose one save folder so the full batch lands in one place." not in response.text
    assert 'id="feedback-modal"' in response.text
    assert 'id="feedback-category"' in response.text
    assert 'aria-label="Batch readiness"' in response.text
    assert 'id="readiness-files-value"' in response.text
    assert 'id="readiness-destination-value"' in response.text
    assert 'id="readiness-conversion-value"' in response.text
    assert 'id="batch-progress-track"' in response.text
    assert 'role="progressbar"' in response.text
    assert 'aria-describedby="batch-progress-detail"' in response.text
    assert 'id="result-banner" role="status"' in response.text
    assert 'id="result-guidance"' in response.text
    assert 'id="result-review"' in response.text
    assert 'id="result-review-list"' in response.text
    assert 'aria-label="Converted PDF review"' in response.text
    assert 'class="timeline-details"' in response.text
    assert "Technical details" in response.text
    assert 'id="result-open-output-button"' in response.text
    assert 'id="result-retry-failed-button"' in response.text
    assert "Retry Failed" in response.text
    assert 'id="result-start-new-button"' in response.text
    assert 'id="filename-style-select"' in response.text
    assert 'id="filename-style-example-value"' in response.text
    assert 'data-testid="dropzone"' in response.text
    assert 'data-testid="file-input"' in response.text
    assert 'data-testid="filename-style-select"' in response.text
    assert 'data-testid="filename-style-example"' in response.text
    assert 'data-testid="conversion-queue"' in response.text
    assert 'data-testid="queue-list"' in response.text
    assert 'data-testid="save-card"' in response.text
    assert 'data-testid="batch-progress"' in response.text
    assert 'data-testid="convert-button"' in response.text
    assert 'data-testid="result-banner"' in response.text
    assert 'data-testid="result-headline"' in response.text
    assert "Project records: Date + subject" in response.text
    assert "Office files: Subject only" in response.text
    assert "Example output" in response.text
    assert 'class="filename-style-panel"' in response.text
    assert response.text.index('id="dropzone"') < response.text.index('id="filename-style-select"')
    assert response.text.index('id="filename-style-select"') < response.text.index("Conversion Queue")
    save_card_start = response.text.index('class="rail-card save-card"')
    save_card_end = response.text.index('class="readiness-list"', save_card_start)
    assert 'id="filename-style-select"' not in response.text[save_card_start:save_card_end]


def test_web_ui_progress_model_keeps_pdf_render_near_completion() -> None:
    source = WEB_UI_APP_PATH.read_text(encoding="utf-8")
    pipeline_model = re.search(r"pipeline_selected:\s*{(?P<body>.*?)}", source, flags=re.DOTALL)

    assert pipeline_model is not None
    assert re.search(r"floor:\s*7[4-9]", pipeline_model.group("body"))
    assert re.search(r"cap:\s*9[7-9]", pipeline_model.group("body"))
    assert re.search(r"baseRate:\s*(1[5-9]|[2-9]\d)", pipeline_model.group("body"))
    assert "progress?.active && progress.stage === item.stage && progress.percent < 100" in source
    assert "existing?.active && existing.stage === item.stage && existing.percent < 100" in source
    assert "retryFailedItems" in source
    assert "Only failed rows are being converted again." in source
    assert "Use Retry Failed to convert only the failed files again." in source
    assert "function failureExplanation" in source
    assert "This is not a valid Outlook .msg email." in source
    assert "resultPanelOwnsOutputAction = isComplete || isAttention" in source
    assert "elements.saveCard.hidden = isComplete || isAttention" in source
    assert "function renderBatchReview" in source
    assert "function reviewFilenameParts" in source
    assert 'lastIndexOf("__")' in source
    assert "result-review-name-row" in source
    assert "result-review-prefix" in source
    assert "result-review-prefix-value" in source
    assert "result-review-extension" in source
    assert "data-result-reveal-path" in source
    assert "data-result-reveal-label" in source
    assert "function revealOutputFile" in source
    assert "function chooseBrowserOutputFolder" in source
    assert "browserOutputDirectoryHandle" in source
    assert "browserDownloadFallback" in source
    assert "function activateBrowserDownloadFallback" in source
    assert "hosted mode uses normal browser downloads instead" in source
    assert "window.showDirectoryPicker" not in source
    assert "state.browserDownloadFallback = isHostedBrowserOutputMode()" in source
    assert "const canChooseFolder = state.capabilities.nativeOutputPicker" in source
    assert "hosted-browser-downloads-3" in (WEB_UI_DIR / "index.html").read_text(encoding="utf-8")
    assert "function downloadOutputFile" in source
    assert "function triggerOutputItemDownload" in source
    assert "sent to browser downloads" in source
    assert "function isTemporaryStorageError" in source
    assert "Temporary conversion storage could not be written to." in source
    assert "/api/output-file/" in source
    assert "function setRevealButtonFeedback" in source
    assert '"/api/reveal-output-file"' in source
    assert 'class="result-review-row is-${isSaved ? "saved" : "retry"}"' in source
    assert "function buildQueueTerminalSummary" in source
    assert "data-queue-details-toggle" in source
    assert "state.queueDetailsExpanded" in source
    assert "visibleItems = shouldCollapseTerminalQueue" in source


def test_web_ui_responsive_breakpoints_prioritize_single_column_layout() -> None:
    css = WEB_UI_CSS_PATH.read_text(encoding="utf-8")

    assert "@media (max-width: 1180px)" in css
    assert re.search(
        r"@media \(max-width: 1180px\).*?\.app-shell\s*{\s*grid-template-columns:\s*minmax\(0,\s*1fr\);",
        css,
        flags=re.DOTALL,
    )
    assert re.search(
        r"@media \(max-width: 1180px\).*?\.ops-rail\s*{\s*grid-template-columns:\s*repeat\(2,\s*minmax\(0,\s*1fr\)\);",
        css,
        flags=re.DOTALL,
    )
    assert re.search(
        r"@media \(max-width: 860px\).*?\.queue-item\s*{\s*grid-template-columns:\s*minmax\(0,\s*1fr\)\s*minmax\(0,\s*1fr\);",
        css,
        flags=re.DOTALL,
    )
    assert re.search(
        r"@media \(max-width: 720px\).*?\.queue-item\s*{\s*grid-template-columns:\s*minmax\(0,\s*1fr\);",
        css,
        flags=re.DOTALL,
    )
    assert ".filename-style-panel" in css
    assert ".result-review-name-row" in css
    assert ".result-review-prefix-value" in css
    assert ".result-review-reveal" in css
    assert ".result-review-reveal.is-confirmed" in css
    assert "-webkit-line-clamp: 2" in css
    assert "overflow-wrap: anywhere" in css
    assert "max-width: 138px" in css
    assert re.search(
        r"@media \(max-width: 860px\).*?\.filename-style-panel\s*{\s*grid-template-columns:\s*minmax\(0,\s*1fr\);",
        css,
        flags=re.DOTALL,
    )


def test_web_ui_animation_performance_budgets_are_source_enforced() -> None:
    css = WEB_UI_CSS_PATH.read_text(encoding="utf-8")
    controller_source = WEB_UI_DROPZONE_CONTROLLER_PATH.read_text(encoding="utf-8")
    pulse_match = re.search(
        r"@keyframes progress-fill-pulse\s*{(?P<body>.*?)\n}\s*\n@keyframes queue-active-glow",
        css,
        flags=re.DOTALL,
    )

    assert pulse_match is not None
    assert "filter:" not in pulse_match.group("body")
    assert "box-shadow:" not in pulse_match.group("body")
    assert "filter: saturate(1.08)" not in css
    assert "filter: blur(1px)" not in css
    assert "contain: layout paint style;" in css
    assert css.count("contain: paint;") >= 4
    assert re.search(
        r"@media \(prefers-reduced-motion: reduce\).*?\.drop-ripple-canvas\s*{\s*display:\s*none;",
        css,
        flags=re.DOTALL,
    )
    assert "RIPPLE_PIXEL_RATIO_CAP = 1.35" in controller_source
    assert "RIPPLE_MAX_CANVAS_AREA = 520000" in controller_source
    assert "RIPPLE_TARGET_FRAME_MS = 1000 / 30" in controller_source
    assert 'window.matchMedia?.("(prefers-reduced-motion: reduce)")?.matches' in controller_source


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


def test_retry_failed_conversion_preserves_completed_items(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    output_dir = tmp_path / "output"
    convert_calls: list[list[str]] = []

    def fake_convert(msg_paths, output_dir_arg, event_sink=None, task_ids_by_source_path=None, **_kwargs) -> ConversionResult:
        assert output_dir_arg == output_dir
        convert_calls.append([path.name for path in msg_paths])
        converted_files = []
        skipped_files = []
        errors = []
        is_first_call = len(convert_calls) == 1
        for msg_path in msg_paths:
            task_id = task_ids_by_source_path[msg_path.resolve()]
            if is_first_call and "bad.msg" in msg_path.name:
                skipped_files.append(msg_path)
                errors.append("Failed to convert bad.msg: simulated parse failed")
                emit_task_event(
                    event_sink,
                    task_id=task_id,
                    stage="failed",
                    file_name=msg_path.name,
                    success=False,
                    error="simulated parse failed",
                )
                continue

            output_path = output_dir_arg / f"{Path(msg_path).stem}.pdf"
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text("pdf", encoding="utf-8")
            converted_files.append(output_path)
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="complete",
                file_name=msg_path.name,
                success=True,
                meta={"outputPath": str(output_path), "outputName": output_path.name},
            )

        return ConversionResult(
            requested_count=len(msg_paths),
            converted_files=converted_files,
            skipped_files=skipped_files,
            errors=errors,
            total_seconds=0.25,
        )

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.convert_msg_files", fake_convert)
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[
            ("files", ("good.msg", BytesIO(b"good"), "application/vnd.ms-outlook")),
            ("files", ("bad.msg", BytesIO(b"bad"), "application/vnd.ms-outlook")),
        ],
    )
    assert upload.status_code == 200
    queued_items = upload.json()["items"]

    first_convert = client.post(
        "/api/convert",
        json={"ids": [item["id"] for item in queued_items], "output_dir": str(output_dir)},
    )
    assert first_convert.status_code == 200
    assert len(first_convert.json()["convertedFiles"]) == 1
    assert first_convert.json()["errors"] == ["Failed to convert bad.msg: simulated parse failed"]

    after_first = client.get("/api/queue").json()["items"]
    assert [item["stage"] for item in after_first] == ["complete", "failed"]
    assert after_first[1]["error"] == "simulated parse failed"

    retry = client.post(
        "/api/convert",
        json={"ids": [item["id"] for item in queued_items], "output_dir": str(output_dir)},
    )
    assert retry.status_code == 200
    assert len(retry.json()["convertedFiles"]) == 1
    assert retry.json()["errors"] == []
    assert len(convert_calls) == 2
    assert any("good.msg" in name for name in convert_calls[0])
    assert any("bad.msg" in name for name in convert_calls[0])
    assert not any("good.msg" in name for name in convert_calls[1])
    assert any("bad.msg" in name for name in convert_calls[1])

    after_retry = client.get("/api/queue").json()["items"]
    assert [item["stage"] for item in after_retry] == ["complete", "complete"]
    assert "error" not in after_retry[1]


def test_twelve_email_batch_stress_preview_progress_failure_and_retry(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    conversion_calls: list[list[str]] = []
    first_attempt_failures = ("batch_03.msg", "batch_09.msg")

    def drain_events(subscriber) -> list[dict[str, object]]:
        payloads = []
        while not subscriber.empty():
            payloads.append(json.loads(subscriber.get_nowait()))
        return payloads

    def fake_parse_msg_file(path: Path) -> EmailRecord:
        index = int(path.stem.split("_")[-1])
        return EmailRecord(
            source_path=path,
            subject=f"Project {index:02d} Update",
            sent_at=datetime(2026, 5, 7, 18 - index, tzinfo=timezone.utc),
            sender=f"Sender {index:02d}",
            to="team@example.com",
            cc="",
            body="Body",
            html_body="",
            attachment_names=[],
            thread_key=normalize_thread_subject(f"Project {index:02d} Update"),
        )

    def fake_convert(msg_paths, output_dir_arg, event_sink=None, task_ids_by_source_path=None, **_kwargs) -> ConversionResult:
        conversion_calls.append([path.name for path in msg_paths])
        converted_files = []
        skipped_files = []
        errors = []
        is_first_attempt = len(conversion_calls) == 1
        for msg_path in msg_paths:
            task_id = task_ids_by_source_path[msg_path.resolve()]
            output_name = f"{msg_path.stem}.pdf"
            output_path = output_dir_arg / output_name
            for stage in ("parse_started", "filename_built", "pdf_pipeline_started", "pipeline_selected", "pdf_written", "deliver_started"):
                emit_task_event(
                    event_sink,
                    task_id=task_id,
                    stage=stage,
                    file_name=msg_path.name,
                    pipeline="stress-html",
                    meta={"outputName": output_name},
                )
            if is_first_attempt and msg_path.name.endswith(first_attempt_failures):
                skipped_files.append(msg_path)
                errors.append(f"Failed to convert {msg_path.name}: simulated stress failure")
                emit_task_event(
                    event_sink,
                    task_id=task_id,
                    stage="failed",
                    file_name=msg_path.name,
                    pipeline="stress-html",
                    success=False,
                    error="simulated stress failure",
                    meta={"outputName": output_name},
                )
                continue

            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text("pdf", encoding="utf-8")
            converted_files.append(output_path)
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="complete",
                file_name=msg_path.name,
                pipeline="stress-html",
                success=True,
                meta={"outputName": output_name, "outputPath": str(output_path)},
            )

        return ConversionResult(
            requested_count=len(msg_paths),
            converted_files=converted_files,
            skipped_files=skipped_files,
            errors=errors,
            total_seconds=3.4,
            timing_lines=["Total 3.40s"],
        )

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.parse_msg_file", fake_parse_msg_file)
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.convert_msg_files", fake_convert)
    app = create_app()
    subscriber = app.state.event_broker.subscribe()
    client = TestClient(app)

    upload_files = [
        ("files", (f"batch_{index:02d}.msg", BytesIO(f"message {index}".encode()), "application/vnd.ms-outlook"))
        for index in range(1, 13)
    ]
    upload = client.post("/api/upload", data={"filename_style": "date_subject"}, files=upload_files)
    assert upload.status_code == 200
    upload_payload = upload.json()
    queued_items = upload_payload["items"]
    assert len(upload_payload["accepted"]) == 12
    assert len(queued_items) == 12
    assert queued_items[0]["outputName"].startswith("2026-05-07_Project 01 Update")

    preview = client.post("/api/filename-style-preview", json={"filename_style": "sender_subject"})
    assert preview.status_code == 200
    preview_items = preview.json()["items"]
    assert len(preview_items) == 12
    assert preview_items[0]["outputName"] == "Sender 01_Project 01 Update.pdf"
    assert preview_items[-1]["outputName"] == "Sender 12_Project 12 Update.pdf"

    first_convert = client.post(
        "/api/convert",
        json={
            "ids": [item["id"] for item in preview_items],
            "output_dir": str(tmp_path / "output"),
            "filename_style": "sender_subject",
        },
    )
    assert first_convert.status_code == 200
    assert len(first_convert.json()["convertedFiles"]) == 10
    assert len(first_convert.json()["errors"]) == 2

    after_first = client.get("/api/queue").json()["items"]
    assert len(after_first) == 12
    assert sum(item["stage"] == "complete" for item in after_first) == 10
    assert sum(item["stage"] == "failed" for item in after_first) == 2
    assert [item["name"] for item in after_first if item["stage"] == "failed"] == ["batch_03.msg", "batch_09.msg"]

    first_events = drain_events(subscriber)
    task_events = [event for event in first_events if "stage" in event]
    assert sum(event["stage"] == "drop_received" for event in task_events) == 12
    assert sum(event["stage"] == "output_folder_selected" for event in task_events) == 12
    assert sum(event["stage"] == "pipeline_selected" for event in task_events) == 12
    assert sum(event["stage"] == "complete" for event in task_events) == 10
    assert sum(event["stage"] == "failed" for event in task_events) == 2

    retry = client.post(
        "/api/convert",
        json={
            "ids": [item["id"] for item in after_first],
            "output_dir": str(tmp_path / "output"),
            "filename_style": "sender_subject",
        },
    )
    assert retry.status_code == 200
    assert len(retry.json()["convertedFiles"]) == 2
    assert retry.json()["errors"] == []
    assert len(conversion_calls) == 2
    assert len(conversion_calls[0]) == 12
    assert [name[-12:] for name in conversion_calls[1]] == ["batch_03.msg", "batch_09.msg"]

    after_retry = client.get("/api/queue").json()["items"]
    assert len(after_retry) == 12
    assert all(item["stage"] == "complete" for item in after_retry)
    assert all(item.get("outputPath") for item in after_retry)


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


def test_upload_previews_thread_date_prefixed_output_names(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")

    def fake_parse_msg_file(path: Path) -> EmailRecord:
        if path.name.endswith("older.msg"):
            subject = "RE: Project Update"
            sent_at = datetime(2026, 4, 10, 18, tzinfo=timezone.utc)
        else:
            subject = "Project Update"
            sent_at = datetime(2026, 4, 30, 18, tzinfo=timezone.utc)
        return EmailRecord(
            source_path=path,
            subject=subject,
            sent_at=sent_at,
            sender="sender@example.com",
            to="to@example.com",
            cc="",
            body="Body",
            html_body="",
            attachment_names=[],
            thread_key=normalize_thread_subject(subject),
        )

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.parse_msg_file", fake_parse_msg_file)
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[
            ("files", ("older.msg", BytesIO(b"older-msg"), "application/vnd.ms-outlook")),
            ("files", ("latest.msg", BytesIO(b"latest-msg"), "application/vnd.ms-outlook")),
        ],
    )

    assert upload.status_code == 200
    items = upload.json()["items"]
    assert [item["outputName"] for item in items] == [
        "2026-04-30_RE_ Project Update.pdf",
        "2026-04-30_Project Update.pdf",
    ]


def test_filename_style_preview_refreshes_queued_output_names(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")

    def fake_parse_msg_file(path: Path) -> EmailRecord:
        return EmailRecord(
            source_path=path,
            subject="Policy Update",
            sent_at=datetime(2026, 4, 30, 18, tzinfo=timezone.utc),
            sender="Jane Smith",
            to="to@example.com",
            cc="",
            body="Body",
            html_body="",
            attachment_names=[],
            thread_key=normalize_thread_subject("Policy Update"),
        )

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.parse_msg_file", fake_parse_msg_file)
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[("files", ("policy.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
    )
    assert upload.status_code == 200
    assert upload.json()["items"][0]["outputName"] == "2026-04-30_Policy Update.pdf"

    preview = client.post("/api/filename-style-preview", json={"filename_style": "sender_subject"})

    assert preview.status_code == 200
    assert preview.json()["filenameStyle"] == "sender_subject"
    assert preview.json()["items"][0]["outputName"] == "Jane Smith_Policy Update.pdf"


def test_convert_forwards_filename_style_to_converter(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    converted_path = tmp_path / "output" / "Jane Smith_Policy Update.pdf"
    seen_filename_style = ""

    def fake_convert(
        msg_paths,
        output_dir,
        event_sink=None,
        task_ids_by_source_path=None,
        filename_style=None,
        **_kwargs,
    ) -> ConversionResult:
        nonlocal seen_filename_style
        seen_filename_style = filename_style
        converted_path.parent.mkdir(parents=True, exist_ok=True)
        converted_path.write_text("pdf", encoding="utf-8")
        task_id = task_ids_by_source_path[msg_paths[0].resolve()]
        emit_task_event(
            event_sink,
            task_id=task_id,
            stage="complete",
            file_name=msg_paths[0].name,
            success=True,
            meta={"outputPath": str(converted_path), "outputName": converted_path.name},
        )
        return ConversionResult(requested_count=1, converted_files=[converted_path], total_seconds=0.25)

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.convert_msg_files", fake_convert)
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[("files", ("policy.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
    )
    queued_item = upload.json()["items"][0]
    convert = client.post(
        "/api/convert",
        json={
            "ids": [queued_item["id"]],
            "output_dir": str(tmp_path / "output"),
            "filename_style": "sender_subject",
        },
    )

    assert convert.status_code == 200
    assert seen_filename_style == "sender_subject"


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


def test_open_output_folder_endpoint_invokes_helper(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    opened: list[str] = []

    def fake_open_output_directory(path_value: str) -> bool:
        opened.append(path_value)
        return True

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.open_output_directory", fake_open_output_directory)
    client = TestClient(create_app())

    response = client.post("/api/open-output-folder", json={"output_dir": str(tmp_path)})

    assert response.status_code == 200
    assert response.json()["ok"] is True
    assert opened == [str(tmp_path)]


def test_reveal_output_file_endpoint_invokes_helper(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    output_path = tmp_path / "output" / "saved.pdf"
    revealed: list[str] = []

    def fake_reveal_output_file(path_value: str) -> bool:
        revealed.append(path_value)
        return True

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.reveal_output_file", fake_reveal_output_file)
    client = TestClient(create_app())

    response = client.post("/api/reveal-output-file", json={"output_path": str(output_path)})

    assert response.status_code == 200
    assert response.json()["ok"] is True
    assert revealed == [str(output_path)]


def test_output_file_endpoint_serves_completed_pdf(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")
    converted_path = tmp_path / "output" / "saved.pdf"

    def fake_convert(msg_paths, output_dir, event_sink=None, task_ids_by_source_path=None, **_kwargs) -> ConversionResult:
        converted_path.parent.mkdir(parents=True, exist_ok=True)
        converted_path.write_bytes(b"%PDF-test")
        task_id = task_ids_by_source_path[msg_paths[0].resolve()]
        emit_task_event(
            event_sink,
            task_id=task_id,
            stage="complete",
            file_name=msg_paths[0].name,
            success=True,
            meta={"outputPath": str(converted_path), "outputName": converted_path.name},
        )
        return ConversionResult(requested_count=1, converted_files=[converted_path], total_seconds=0.25)

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.convert_msg_files", fake_convert)
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[("files", ("sample.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
    )
    queued_item = upload.json()["items"][0]

    convert = client.post(
        "/api/convert",
        json={"ids": [queued_item["id"]], "output_dir": str(tmp_path / "output")},
    )
    assert convert.status_code == 200

    response = client.get(f"/api/output-file/{queued_item['id']}")

    assert response.status_code == 200
    assert response.headers["content-type"].startswith("application/pdf")
    assert response.content == b"%PDF-test"


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
