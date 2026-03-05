from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import msg_to_pdf_dropzone.corpus_profiler as profiler
from msg_to_pdf_dropzone.converter import ConversionResult, FileTimingRecord
from msg_to_pdf_dropzone.models import EmailRecord
from msg_to_pdf_dropzone.thread_logic import normalize_thread_subject


def _fake_record(path: Path) -> EmailRecord:
    return EmailRecord(
        source_path=path,
        subject=f"Subject {path.stem}",
        sent_at=datetime(2026, 3, 1, 12, 0, tzinfo=timezone.utc),
        sender="sender@example.com",
        to="to@example.com",
        cc="",
        body="Body",
        html_body="<html><body><p>Body</p></body></html>",
        attachment_names=[],
        thread_key=normalize_thread_subject(f"Subject {path.stem}"),
    )


def test_profile_corpus_writes_json_and_markdown(monkeypatch, tmp_path: Path) -> None:
    emails_dir = tmp_path / "emails"
    emails_dir.mkdir()
    msg_one = emails_dir / "one.msg"
    msg_two = emails_dir / "two.msg"
    msg_one.write_text("dummy", encoding="utf-8")
    msg_two.write_text("dummy", encoding="utf-8")

    monkeypatch.setattr(profiler, "parse_msg_file", lambda path: _fake_record(path))

    def fake_convert(paths: list[Path], output_dir: Path) -> ConversionResult:
        output_file = output_dir / f"{paths[0].stem}.pdf"
        output_file.write_bytes(b"%PDF-1.4\nfake\n")
        result = ConversionResult(requested_count=1)
        result.converted_files.append(output_file)
        result.parse_seconds = 0.05
        result.write_seconds = 0.25
        result.total_seconds = 0.30
        result.file_timing_records.append(
            FileTimingRecord(
                file_name=paths[0].name,
                parse_seconds=0.05,
                filename_seconds=0.0,
                pdf_seconds=0.25,
                total_seconds=0.30,
                pipeline="edge_html",
                success=True,
                stage_seconds={"edge_html": 0.25},
                image_metrics={
                    "total_images": 3,
                    "cid_resolved": 2,
                    "cid_unresolved": 0,
                    "signature_small_dropped": 1,
                    "remote_dropped": 1,
                },
            )
        )
        result.timing_lines = [
            "Total 0.30s (parse 0.05s, thread 0.00s, pdf 0.25s)",
            f"{paths[0].name}: parse 0.05s, filename 0.00s, pdf 0.25s, total 0.30s [pipeline edge_html]",
        ]
        return result

    monkeypatch.setattr(profiler, "convert_msg_files", fake_convert)

    summary, json_path, markdown_path = profiler.profile_corpus(
        emails_dir,
        tmp_path / "reports",
        runs=2,
        render_strategy="fidelity",
    )

    assert summary["run_count"] == 2
    assert summary["file_count"] == 2
    assert summary["acceptance"]["parse_failures"] == 0
    assert summary["acceptance"]["conversion_failures"] == 0

    assert json_path.exists()
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    assert payload["aggregate"]["count"] == 4
    assert payload["aggregate"]["failure_count"] == 0
    assert payload["aggregate"]["image_metrics"]["total_images"]["sum"] == 12
    assert payload["runs"][0]["files"][0]["image_metrics"]["cid_resolved"] == 2

    assert markdown_path.exists()
    markdown_text = markdown_path.read_text(encoding="utf-8")
    assert "MSG Corpus Profiling Report" in markdown_text
    assert "Aggregate Metrics" in markdown_text
    assert "Image Metrics (Aggregate)" in markdown_text


def test_percentile_handles_empty_values() -> None:
    assert profiler._percentile([], 0.95) == 0.0
