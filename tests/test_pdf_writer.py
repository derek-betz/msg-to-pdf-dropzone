from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

import msg_to_pdf_dropzone.pdf_writer as pdf_writer
from msg_to_pdf_dropzone.models import EmailRecord, InlineImageAsset
from msg_to_pdf_dropzone.pdf_writer import PdfWriteDiagnostics, build_email_html_document, write_email_pdf
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
    assert 'class="message-header"' in html_document
    assert 'class="meta"' not in html_document
    assert "Attachments</h3>" not in html_document
    from_index = html_document.index(">From:</div>")
    sent_index = html_document.index(">Sent:</div>")
    to_index = html_document.index(">To:</div>")
    cc_index = html_document.index(">Cc:</div>")
    subject_index = html_document.index(">Subject:</div>")
    attachments_index = html_document.index(">Attachments:</div>")
    assert from_index < sent_index < to_index < cc_index < subject_index < attachments_index


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


def test_write_email_pdf_collects_stage_diagnostics(monkeypatch, tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=tmp_path / "sample.msg",
        subject="Diagnostics Subject",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Plain text body",
        html_body="",
        attachment_names=[],
        thread_key=normalize_thread_subject("Diagnostics Subject"),
    )
    record.source_path.write_text("dummy", encoding="utf-8")

    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_outlook_and_edge", lambda _msg, _out: False)
    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_edge_html", lambda _html, _out: False)

    def fake_reportlab(_record: EmailRecord, output_path: Path) -> None:
        output_path.write_bytes(b"%PDF-1.4\nfake\n")

    monkeypatch.setattr(pdf_writer, "_write_pdf_via_reportlab", fake_reportlab)

    diagnostics = PdfWriteDiagnostics()
    output_file = tmp_path / "diagnostics.pdf"
    write_email_pdf(record, output_file, diagnostics=diagnostics)

    assert output_file.exists()
    assert diagnostics.pipeline == "reportlab"
    assert diagnostics.total_seconds >= 0.0
    assert diagnostics.stage_seconds["outlook_edge"] >= 0.0
    assert diagnostics.stage_seconds["build_html"] >= 0.0
    assert diagnostics.stage_seconds["edge_html"] >= 0.0
    assert diagnostics.stage_seconds["reportlab"] >= 0.0


def test_write_email_pdf_fast_strategy_skips_outlook_stage(monkeypatch, tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=tmp_path / "sample.msg",
        subject="Fast Strategy Subject",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Body text",
        html_body="<html><body><p>Fast body</p></body></html>",
        attachment_names=[],
        thread_key=normalize_thread_subject("Fast Strategy Subject"),
    )
    record.source_path.write_text("dummy", encoding="utf-8")

    monkeypatch.setenv("MSG_TO_PDF_RENDER_STRATEGY", "fast")
    monkeypatch.setattr(
        pdf_writer,
        "_try_write_pdf_via_outlook_and_edge",
        lambda _msg, _out: (_ for _ in ()).throw(AssertionError("Outlook stage should be skipped in fast mode.")),
    )

    def fake_edge_html(_html_document: str, output_path: Path) -> bool:
        output_path.write_bytes(b"%PDF-1.4\nfake-fast\n")
        return True

    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_edge_html", fake_edge_html)

    diagnostics = PdfWriteDiagnostics()
    output_file = tmp_path / "fast.pdf"
    write_email_pdf(record, output_file, diagnostics=diagnostics)

    assert output_file.exists()
    assert diagnostics.pipeline == "edge_html"
    assert diagnostics.stage_seconds["render_strategy_fast"] == 1.0
    assert diagnostics.stage_seconds["outlook_edge"] == 0.0


def test_write_email_pdf_emits_pipeline_selection_events(monkeypatch, tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=tmp_path / "sample.msg",
        subject="Pipeline Events",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Plain text body",
        html_body="",
        attachment_names=[],
        thread_key=normalize_thread_subject("Pipeline Events"),
    )
    record.source_path.write_text("dummy", encoding="utf-8")

    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_outlook_and_edge", lambda _msg, _out: False)
    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_edge_html", lambda _html, _out: False)
    monkeypatch.setattr(
        pdf_writer,
        "_write_pdf_via_reportlab",
        lambda _record, output_path: output_path.write_bytes(b"%PDF-1.4\nfake\n"),
    )

    events = []
    diagnostics = PdfWriteDiagnostics()
    output_file = tmp_path / "events.pdf"
    write_email_pdf(
        record,
        output_file,
        diagnostics=diagnostics,
        event_sink=events.append,
        task_id="task-456",
        event_meta={"batchId": "msg-batch-123", "batchSize": 4, "batchIndex": 1},
    )

    assert output_file.exists()
    assert [event.stage for event in events] == [
        "pipeline_selected",
        "pipeline_selected",
        "pipeline_selected",
    ]
    assert [event.pipeline for event in events] == [
        "outlook_edge",
        "edge_html",
        "reportlab",
    ]
    assert all(event.task_id == "task-456" for event in events)
    assert all(event.meta is not None and event.meta["batchId"] == "msg-batch-123" for event in events)


def test_build_email_html_document_rewrites_cid_and_filters_signature_images() -> None:
    hero_image = InlineImageAsset(
        cid="hero-image",
        mime_type="image/png",
        filename="hero.png",
        data=b"x" * (50 * 1024),
        size_bytes=50 * 1024,
    )
    small_logo = InlineImageAsset(
        cid="small-logo",
        mime_type="image/png",
        filename="logo.png",
        data=b"y" * (6 * 1024),
        size_bytes=6 * 1024,
    )
    record = EmailRecord(
        source_path=Path("sample.msg"),
        subject="Image Test",
        sent_at=datetime(2026, 3, 4, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Plain text fallback",
        html_body=(
            "<html><body>"
            "<p>See attached screenshot:</p>"
            "<img src='cid:hero-image' width='640' height='360' />"
            "<p>Thanks,</p><p>Jane</p>"
            "<img src='cid:small-logo' width='140' height='60' />"
            "<img src='https://example.com/tracker.png' width='80' height='40' />"
            "</body></html>"
        ),
        attachment_names=[],
        thread_key=normalize_thread_subject("Image Test"),
        inline_images=[hero_image, small_logo],
    )

    diagnostics = PdfWriteDiagnostics()
    html_document = build_email_html_document(record, diagnostics=diagnostics)

    assert html_document.count("data:image/png;base64,") == 1
    assert "cid:hero-image" not in html_document
    assert "cid:small-logo" not in html_document
    assert "https://example.com/tracker.png" not in html_document
    assert diagnostics.image_metrics["total_images"] == 3
    assert diagnostics.image_metrics["cid_resolved"] == 1
    assert diagnostics.image_metrics["signature_small_dropped"] == 1
    assert diagnostics.image_metrics["remote_dropped"] == 1
    assert diagnostics.image_metrics["cid_unresolved"] == 0


def test_write_email_pdf_falls_back_to_cid_html_after_outlook_attempt(monkeypatch, tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=tmp_path / "sample.msg",
        subject="Auto CID",
        sent_at=datetime(2026, 3, 4, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Body text",
        html_body="<html><body><p>Hello</p><img src='cid:hero' width='500' height='300' /></body></html>",
        attachment_names=[],
        thread_key=normalize_thread_subject("Auto CID"),
        inline_images=[
            InlineImageAsset(
                cid="hero",
                mime_type="image/png",
                filename="hero.png",
                data=b"z" * (45 * 1024),
                size_bytes=45 * 1024,
            )
        ],
    )
    record.source_path.write_text("dummy", encoding="utf-8")

    monkeypatch.setenv("MSG_TO_PDF_RENDER_STRATEGY", "fidelity")
    attempted = {"outlook": 0}

    def fake_outlook(_msg: Path, _out: Path) -> bool:
        attempted["outlook"] += 1
        return False

    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_outlook_and_edge", fake_outlook)

    captured: dict[str, str] = {}

    def fake_edge_html(html_document: str, output_path: Path) -> bool:
        captured["html"] = html_document
        output_path.write_bytes(b"%PDF-1.4\nfake-cid\n")
        return True

    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_edge_html", fake_edge_html)

    diagnostics = PdfWriteDiagnostics()
    output_file = tmp_path / "cid-auto.pdf"
    write_email_pdf(record, output_file, diagnostics=diagnostics)

    assert output_file.exists()
    assert diagnostics.pipeline == "edge_html"
    assert diagnostics.stage_seconds["auto_cid_html"] == 1.0
    assert attempted["outlook"] == 1
    assert diagnostics.stage_seconds["outlook_edge"] >= 0.0
    assert diagnostics.image_metrics["cid_resolved"] == 1
    assert "data:image/png;base64," in captured["html"]


def test_print_web_document_via_edge_disables_pdf_headers_and_footers(
    monkeypatch,
    tmp_path: Path,
) -> None:
    input_path = tmp_path / "sample.html"
    input_path.write_text("<html><body><p>Hello</p></body></html>", encoding="utf-8")
    output_path = tmp_path / "sample.pdf"

    monkeypatch.setattr(pdf_writer.os, "name", "nt", raising=False)
    monkeypatch.setattr(
        pdf_writer,
        "_find_edge_executable",
        lambda: Path("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"),
    )

    captured: dict[str, object] = {}

    def fake_run(command: list[str], **_kwargs: object) -> SimpleNamespace:
        captured["command"] = command
        output_path.write_bytes(b"%PDF-1.4\nfake\n")
        return SimpleNamespace(returncode=0)

    monkeypatch.setattr(pdf_writer.subprocess, "run", fake_run)

    assert pdf_writer._print_web_document_via_edge(input_path, output_path) is True
    command = captured["command"]
    assert isinstance(command, list)
    assert "--no-pdf-header-footer" in command
    assert any(part.startswith("--print-to-pdf=") for part in command)
