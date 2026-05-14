from __future__ import annotations

import base64
from datetime import datetime, timezone
from pathlib import Path
import threading
from types import SimpleNamespace

import msg_to_pdf_dropzone.pdf_writer as pdf_writer
from msg_to_pdf_dropzone.models import EmailRecord, InlineImageAsset
from msg_to_pdf_dropzone.pdf_writer import (
    PdfWriteDiagnostics,
    _as_paragraph_text,
    _estimate_data_uri_bytes,
    _extract_body_html_fragment,
    _format_body_blocks,
    _format_sent_value,
    _is_content_image,
    _is_small_signature_image,
    _parse_numeric_dimension,
    build_email_html_document,
    write_email_pdf,
)
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


def test_write_email_pdf_prefers_outlook_edge_pipeline_in_fidelity_mode(monkeypatch, tmp_path: Path) -> None:
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
    monkeypatch.setenv("MSG_TO_PDF_RENDER_STRATEGY", "fidelity")

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


def test_write_email_pdf_defaults_to_edge_html_pipeline(monkeypatch, tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=tmp_path / "sample.msg",
        subject="Default Strategy Subject",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Body text",
        html_body="<html><body><p>Default route body</p></body></html>",
        attachment_names=[],
        thread_key=normalize_thread_subject("Default Strategy Subject"),
    )
    record.source_path.write_text("dummy", encoding="utf-8")
    monkeypatch.delenv("MSG_TO_PDF_RENDER_STRATEGY", raising=False)
    monkeypatch.setattr(
        pdf_writer,
        "_try_write_pdf_via_outlook_and_edge",
        lambda _msg, _out: (_ for _ in ()).throw(AssertionError("Outlook stage should be skipped by default.")),
    )

    def fake_edge_html(_html_document: str, output_path: Path) -> bool:
        output_path.write_bytes(b"%PDF-1.4\ndefault-fast\n")
        return True

    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_edge_html", fake_edge_html)

    diagnostics = PdfWriteDiagnostics()
    output_file = tmp_path / "default-fast.pdf"
    write_email_pdf(record, output_file, diagnostics=diagnostics)

    assert output_file.exists()
    assert diagnostics.pipeline == "edge_html"
    assert diagnostics.stage_seconds["render_strategy_fast"] == 1.0
    assert diagnostics.stage_seconds["outlook_edge"] == 0.0


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
    monkeypatch.setenv("MSG_TO_PDF_RENDER_STRATEGY", "fidelity")

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
    monkeypatch.setenv("MSG_TO_PDF_RENDER_STRATEGY", "fidelity")

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


def test_write_email_pdf_falls_back_to_cid_html_after_outlook_attempt_for_small_inline_images(
    monkeypatch,
    tmp_path: Path,
) -> None:
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
                data=b"z" * (12 * 1024),
                size_bytes=12 * 1024,
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


def test_write_email_pdf_prefers_html_first_for_large_inline_images(monkeypatch, tmp_path: Path) -> None:
    record = EmailRecord(
        source_path=tmp_path / "sample.msg",
        subject="Large Inline Image",
        sent_at=datetime(2026, 3, 4, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Body text",
        html_body="<html><body><p>Hello</p><img src='cid:hero' width='800' height='600' /></body></html>",
        attachment_names=[],
        thread_key=normalize_thread_subject("Large Inline Image"),
        inline_images=[
            InlineImageAsset(
                cid="hero",
                mime_type="image/png",
                filename="hero.png",
                data=b"z" * (60 * 1024),
                size_bytes=60 * 1024,
            )
        ],
    )
    record.source_path.write_text("dummy", encoding="utf-8")

    call_order: list[str] = []

    def fake_outlook(_msg: Path, _out: Path) -> bool:
        call_order.append("outlook")
        _out.write_bytes(b"%PDF-1.4\nfake-outlook\n")
        return True

    def fake_edge_html(html_document: str, output_path: Path) -> bool:
        call_order.append("edge_html")
        output_path.write_bytes(b"%PDF-1.4\nfake-edge\n")
        return True

    monkeypatch.setenv("MSG_TO_PDF_RENDER_STRATEGY", "fidelity")
    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_outlook_and_edge", fake_outlook)
    monkeypatch.setattr(pdf_writer, "_try_write_pdf_via_edge_html", fake_edge_html)

    diagnostics = PdfWriteDiagnostics()
    output_file = tmp_path / "large-inline.pdf"
    write_email_pdf(record, output_file, diagnostics=diagnostics)

    assert output_file.exists()
    assert diagnostics.pipeline == "edge_html"
    assert diagnostics.stage_seconds["prefer_html_inline_images"] == 1.0
    assert call_order == ["edge_html"]


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
    assert "--print-to-pdf-no-header" in command
    assert any(part.startswith("--print-to-pdf=") for part in command)


def test_print_web_document_via_edge_waits_for_delayed_output(
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

    def fake_run(_command: list[str], **_kwargs: object) -> SimpleNamespace:
        timer = threading.Timer(0.05, lambda: output_path.write_bytes(b"%PDF-1.4\nfake\n"))
        timer.start()
        return SimpleNamespace(returncode=0)

    monkeypatch.setattr(pdf_writer.subprocess, "run", fake_run)

    assert pdf_writer._print_web_document_via_edge(input_path, output_path) is True


def test_edge_html_pipeline_uses_staging_sibling_temp_dir(
    monkeypatch,
    tmp_path: Path,
) -> None:
    staging_dir = tmp_path / "staging"
    output_path = tmp_path / "output" / "sample.pdf"
    output_path.parent.mkdir()
    seen: dict[str, Path] = {}

    monkeypatch.delenv("MSG_TO_PDF_TEMP_DIR", raising=False)
    monkeypatch.setenv("MSG_TO_PDF_STAGING_DIR", str(staging_dir))
    monkeypatch.setattr(pdf_writer.os, "name", "nt", raising=False)

    def fake_print(input_path: Path, target_path: Path) -> bool:
        seen["input_path"] = input_path
        target_path.write_bytes(b"%PDF-1.4\nfake\n")
        return True

    monkeypatch.setattr(pdf_writer, "_print_web_document_via_edge", fake_print)

    assert pdf_writer._try_write_pdf_via_edge_html("<html><body>Hello</body></html>", output_path) is True
    assert seen["input_path"].parent.parent == tmp_path / "temp"
    assert output_path.exists()
    assert not seen["input_path"].exists()


def test_format_body_blocks_splits_on_blank_lines() -> None:
    result = _format_body_blocks("First paragraph.\n\nSecond paragraph.")
    assert result == ["First paragraph.", "Second paragraph."]


def test_format_body_blocks_returns_placeholder_for_empty_body() -> None:
    assert _format_body_blocks("") == ["(No body text.)"]
    assert _format_body_blocks("   ") == ["(No body text.)"]


def test_format_body_blocks_normalizes_crlf() -> None:
    result = _format_body_blocks("Line one\r\n\r\nLine two")
    assert result == ["Line one", "Line two"]


def test_as_paragraph_text_escapes_html_special_chars() -> None:
    result = _as_paragraph_text("Hello <world> & 'you'")
    assert "&lt;world&gt;" in result
    assert "&amp;" in result


def test_as_paragraph_text_replaces_newlines_with_br() -> None:
    result = _as_paragraph_text("line one\nline two")
    assert "<br/>" in result


def test_extract_body_html_fragment_extracts_body_content() -> None:
    html = "<html><body><p>Hello world</p></body></html>"
    result = _extract_body_html_fragment(html)
    assert "<p>Hello world</p>" in result


def test_extract_body_html_fragment_removes_script_tags() -> None:
    html = "<html><body><script>evil()</script><p>Safe</p></body></html>"
    result = _extract_body_html_fragment(html)
    assert "evil()" not in result
    assert "<p>Safe</p>" in result


def test_extract_body_html_fragment_returns_empty_for_empty_input() -> None:
    assert _extract_body_html_fragment("") == ""
    assert _extract_body_html_fragment("   ") == ""


def test_extract_body_html_fragment_returns_full_content_when_no_body_tag() -> None:
    html = "<p>No body tag here</p>"
    result = _extract_body_html_fragment(html)
    assert "<p>No body tag here</p>" in result


def test_parse_numeric_dimension_returns_integer_from_plain_number() -> None:
    assert _parse_numeric_dimension("100") == 100


def test_parse_numeric_dimension_extracts_number_from_css_value() -> None:
    # Python banker's rounding: round(100.5) == 100, round(101.5) == 102
    assert _parse_numeric_dimension("100.5px") == 100
    assert _parse_numeric_dimension("101.5px") == 102
    assert _parse_numeric_dimension("200px") == 200


def test_parse_numeric_dimension_returns_none_for_zero() -> None:
    assert _parse_numeric_dimension("0") is None


def test_parse_numeric_dimension_returns_none_for_empty() -> None:
    assert _parse_numeric_dimension("") is None


def test_parse_numeric_dimension_returns_none_for_non_numeric() -> None:
    assert _parse_numeric_dimension("auto") is None


def test_is_small_signature_image_with_small_dims_and_size() -> None:
    assert _is_small_signature_image(100, 50, 5 * 1024) is True


def test_is_small_signature_image_with_large_dimension() -> None:
    assert _is_small_signature_image(300, 200, 5 * 1024) is False


def test_is_small_signature_image_with_large_byte_size() -> None:
    assert _is_small_signature_image(100, 50, 25 * 1024) is False


def test_is_small_signature_image_with_no_dimensions() -> None:
    assert _is_small_signature_image(None, None, 5 * 1024) is False


def test_is_content_image_with_large_file_size() -> None:
    assert _is_content_image(100, 50, 45 * 1024) is True


def test_is_content_image_with_large_dimension() -> None:
    assert _is_content_image(300, 100, 1024) is True


def test_is_content_image_with_large_area() -> None:
    assert _is_content_image(250, 200, 1024) is True


def test_is_content_image_returns_false_for_small_image() -> None:
    assert _is_content_image(100, 50, 1024) is False


def test_estimate_data_uri_bytes_base64_encoded() -> None:
    data = b"hello world test data payload"
    encoded = base64.b64encode(data).decode("ascii")
    uri = f"data:image/png;base64,{encoded}"
    estimated = _estimate_data_uri_bytes(uri)
    assert abs(estimated - len(data)) <= 3


def test_estimate_data_uri_bytes_returns_zero_for_missing_comma() -> None:
    assert _estimate_data_uri_bytes("no comma here") == 0


def test_format_sent_value_formats_utc_datetime() -> None:
    dt = datetime(2026, 3, 15, 10, 30, 0, tzinfo=timezone.utc)
    result = _format_sent_value(dt)
    expected = dt.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z").strip()
    assert result == expected


def test_build_email_html_document_renders_plain_body_when_no_html_body() -> None:
    record = EmailRecord(
        source_path=Path("sample.msg"),
        subject="Plain Body Test",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="This is plain text.",
        html_body="",
        attachment_names=[],
        thread_key=normalize_thread_subject("Plain Body Test"),
    )
    html = build_email_html_document(record)
    assert "plain-body" in html
    assert "This is plain text." in html


def test_build_email_html_document_escapes_special_chars_in_subject() -> None:
    record = EmailRecord(
        source_path=Path("sample.msg"),
        subject="Subject <with> & special",
        sent_at=datetime(2026, 2, 20, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Body",
        html_body="",
        attachment_names=[],
        thread_key=normalize_thread_subject("Subject <with> & special"),
    )
    html = build_email_html_document(record)
    assert "&lt;with&gt;" in html
    assert "&amp;" in html
