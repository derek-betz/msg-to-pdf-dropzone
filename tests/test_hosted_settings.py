from __future__ import annotations

from io import BytesIO
from pathlib import Path

from fastapi.testclient import TestClient

from msg_to_pdf_dropzone.converter import ConversionResult
from msg_to_pdf_dropzone.task_events import emit_task_event
from msg_to_pdf_dropzone.web_server import create_app


def test_settings_exposes_server_managed_output_dir(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_OUTPUT_DIR", str(tmp_path / "pdf-output"))
    monkeypatch.setenv("MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT", "1")
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")

    client = TestClient(create_app())

    settings = client.get("/api/settings")

    assert settings.status_code == 200
    payload = settings.json()
    assert payload["serverMode"] is True
    assert payload["defaultOutputDir"] == str(tmp_path / "pdf-output")
    assert payload["capabilities"]["nativeOutputPicker"] is False
    assert payload["capabilities"]["outlookImport"] is False


def test_choose_output_folder_returns_server_managed_dir(monkeypatch, tmp_path: Path) -> None:
    output_dir = tmp_path / "pdf-output"
    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_OUTPUT_DIR", str(output_dir))
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")

    client = TestClient(create_app())

    response = client.post("/api/choose-output-folder")

    assert response.status_code == 200
    assert response.json()["outputDir"] == str(output_dir)
    assert output_dir.exists()


def test_convert_uses_default_output_dir_when_request_omits_one(monkeypatch, tmp_path: Path) -> None:
    output_dir = tmp_path / "pdf-output"
    converted_path = output_dir / "sample.pdf"

    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_OUTPUT_DIR", str(output_dir))
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")

    def fake_convert(msg_paths, output_dir, event_sink=None, task_ids_by_source_path=None, **_kwargs) -> ConversionResult:
        assert output_dir == tmp_path / "pdf-output"
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
            total_seconds=0.2,
            timing_lines=["Total 0.20s"],
        )

    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.convert_msg_files", fake_convert)
    client = TestClient(create_app())

    upload = client.post(
        "/api/upload",
        files=[("files", ("sample.msg", BytesIO(b"msg-bytes"), "application/vnd.ms-outlook"))],
    )

    queued_item = upload.json()["items"][0]
    convert = client.post("/api/convert", json={"ids": [queued_item["id"]]})

    assert convert.status_code == 200
    assert convert.json()["convertedFiles"] == [str(converted_path)]


def test_import_outlook_can_be_disabled_for_hosted_mode(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT", "1")
    monkeypatch.setattr("msg_to_pdf_dropzone.web_server.STAGING_DIR", tmp_path / "staging")

    client = TestClient(create_app())

    response = client.post("/api/import-outlook")

    assert response.status_code == 403
