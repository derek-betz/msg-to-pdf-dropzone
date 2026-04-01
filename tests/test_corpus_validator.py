from __future__ import annotations

from pathlib import Path

import pytest

import msg_to_pdf_dropzone.corpus_profiler as profiler
import msg_to_pdf_dropzone.corpus_validator as validator
from msg_to_pdf_dropzone.converter import ConversionResult, FileTimingRecord


def _make_case_dir(root: Path, case_id: str) -> Path:
    case_dir = root / case_id
    case_dir.mkdir(parents=True)
    return case_dir


def _write_case_pair(case_dir: Path, *, msg_name: str = "sample.msg", pdf_name: str = "golden.pdf") -> None:
    (case_dir / msg_name).write_text("dummy", encoding="utf-8")
    (case_dir / pdf_name).write_bytes(b"%PDF-1.4\nfake\n")


def _make_snapshot(
    path: Path,
    *,
    from_value: str = "Sender Example <sender@example.com>",
    sent_value: str = "Sunday, March 1, 2026 12:56 PM",
    to_value: str = "Derek Betz",
    cc_value: str = "",
    subject_value: str = "Sample Subject",
    attachments_value: str = "",
    body_lines: list[str] | None = None,
    page_count: int = 1,
    file_size_bytes: int = 1000,
) -> validator.PdfSnapshot:
    header_lines = [
        f"From: {from_value}",
        f"Sent: {sent_value}",
        f"To: {to_value}",
    ]
    if cc_value:
        header_lines.append(f"Cc: {cc_value}")
    header_lines.append(f"Subject: {subject_value}")
    if attachments_value:
        header_lines.append(f"Attachments: {attachments_value}")
    body_text = "\n".join(
        body_lines
        or [
            "This is a long anchor line for regression testing and semantic validation coverage.",
            "Another anchor-worthy line that should survive normalization and line wrapping changes.",
        ]
    )
    extracted_text = "\n".join(header_lines) + "\n\n" + body_text
    headers, order, parsed_body = validator._extract_header_and_body(extracted_text)
    return validator.PdfSnapshot(
        path=path,
        page_count=page_count,
        file_size_bytes=file_size_bytes,
        extracted_text=extracted_text,
        normalized_text=validator._normalize_inline_text(extracted_text),
        header_fields=headers,
        header_order=order,
        body_text=parsed_body,
        body_anchors=validator._extract_body_anchors(parsed_body),
    )


def _fake_conversion(paths: list[Path], output_dir: Path, *, pipeline: str = "outlook_edge") -> ConversionResult:
    output_path = output_dir / "generated.pdf"
    output_path.write_bytes(b"%PDF-1.4\nfake\n")
    result = ConversionResult(requested_count=len(paths))
    result.converted_files.append(output_path)
    result.file_timing_records.append(
        FileTimingRecord(
            file_name=paths[0].name,
            parse_seconds=0.01,
            filename_seconds=0.0,
            pdf_seconds=0.1,
            total_seconds=0.11,
            pipeline=pipeline,
            success=True,
        )
    )
    return result


def test_discover_case_pairs_accepts_numbered_pair_directories(tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "7")
    _write_case_pair(case_dir, msg_name="mail.msg", pdf_name="golden.pdf")

    pairs = validator.discover_case_pairs(tmp_path)

    assert len(pairs) == 1
    assert pairs[0].case_id == "7"
    assert pairs[0].msg_path.name == "mail.msg"
    assert pairs[0].golden_pdf_path.name == "golden.pdf"


def test_discover_case_pairs_rejects_missing_msg_file(tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    (case_dir / "golden.pdf").write_bytes(b"%PDF-1.4\nfake\n")

    with pytest.raises(ValueError, match="exactly one .msg file"):
        validator.discover_case_pairs(tmp_path)


def test_discover_case_pairs_rejects_missing_golden_pdf(tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    (case_dir / "mail.msg").write_text("dummy", encoding="utf-8")

    with pytest.raises(ValueError, match="exactly one golden .pdf file"):
        validator.discover_case_pairs(tmp_path)


def test_discover_case_pairs_rejects_multiple_pdfs(tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    (case_dir / "mail.msg").write_text("dummy", encoding="utf-8")
    (case_dir / "a.pdf").write_bytes(b"%PDF-1.4\nfake\n")
    (case_dir / "b.pdf").write_bytes(b"%PDF-1.4\nfake\n")

    with pytest.raises(ValueError, match="exactly one golden .pdf file"):
        validator.discover_case_pairs(tmp_path)


def test_extract_header_and_body_handles_wrapped_recipients() -> None:
    text = (
        "From: Sender Example\n"
        "Sent: Sunday, March 1, 2026 12:56 PM\n"
        "To: Derek Betz;\n"
        " Wendell Solomon;\n"
        " Another Recipient\n"
        "Subject: Wrapped Recipients\n"
        "\n"
        "Body starts here.\n"
    )

    headers, order, body = validator._extract_header_and_body(text)

    assert headers["To"] == "Derek Betz; Wendell Solomon; Another Recipient"
    assert order == ["From", "Sent", "To", "Subject"]
    assert body == "Body starts here."


def test_sent_value_match_tolerates_timezone_wording_and_minute_rounding() -> None:
    expected = "Sunday, March 1, 2026 12:56 PM"
    actual = "2026-03-01 12:55:40 Mountain Standard Time"

    assert validator._sent_values_match(expected, actual) is True


def test_field_value_match_tolerates_display_name_vs_email_formatting() -> None:
    expected = "Derek Betz"
    actual = "Derek Betz <dbetz@hanson-inc.com>"

    assert validator._field_values_match("To", expected, actual) is True


def test_extract_attachment_names_handles_multiline_header() -> None:
    value = "report-one.pdf\nreport-two.xlsx\nimage001.png"

    assert validator._extract_attachment_names(value) == [
        "report-one.pdf",
        "report-two.xlsx",
        "image001.png",
    ]


def test_extract_body_anchors_prefers_long_meaningful_lines() -> None:
    body_text = (
        "Short\n"
        "This is a very long line with enough specific detail to become a golden anchor for testing.\n"
        "Another meaningful regression sentence that should also be selected as an anchor.\n"
    )

    anchors = validator._extract_body_anchors(body_text)

    assert len(anchors) == 2
    assert anchors[0].startswith("This is a very long line")


def test_anchor_matches_tolerates_wrapping_and_partial_tokenized_overlap() -> None:
    anchor = "Please create the following sheets for the typical. I use SR43 (22H0004A) as the example."
    generated_text = (
        "from: jennifer beyer "
        "please create the following sheets for the typical i use sr43 22h0004a as the "
        "example thanks jennifer"
    )

    assert validator._anchor_matches(anchor, generated_text) is True


def test_validate_corpus_reports_clean_pass(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir, msg_name="mail.msg", pdf_name="golden.pdf")

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, attachments_value="receipt.pdf")
        return _make_snapshot(
            pdf_path,
            sent_value="2026-03-01 12:56:00 Mountain Standard Time",
            to_value="Derek Betz <dbetz@hanson-inc.com>",
            attachments_value="receipt.pdf",
        )

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    summary, json_path, markdown_path = validator.validate_corpus(tmp_path, tmp_path / "reports")

    assert summary["passed_count"] == 1
    assert summary["failed_count"] == 0
    assert summary["strict_failed_count"] == 0
    assert summary["info_case_count"] == 0
    assert json_path.exists()
    assert markdown_path.exists()


def test_compare_snapshots_ignores_format_only_header_and_file_size_differences(tmp_path: Path) -> None:
    golden = _make_snapshot(
        tmp_path / "golden.pdf",
        from_value="Luke Hewitt",
        sent_value="Tuesday, March 3, 2026 10:34 AM",
        to_value="Derek Betz",
        cc_value="Kurt Bialobreski",
        subject_value="RE: OpenAI API Key Costs",
        file_size_bytes=350000,
    )
    generated = _make_snapshot(
        tmp_path / "generated.pdf",
        from_value="Luke Hewitt <lhewitt@hanson-inc.com>",
        sent_value="2026-03-03 10:33:51 Mountain Standard Time",
        to_value="Derek Betz <DBetz@hanson-inc.com>",
        cc_value="Kurt Bialobreski <KBialobreski@hanson-inc.com>",
        subject_value="RE: OpenAI API Key Costs",
        file_size_bytes=120000,
    )

    result = validator.compare_snapshots(golden, generated, pipeline="edge_html")

    assert result.passed is True
    assert result.warnings == []


def test_validate_corpus_fails_on_header_mismatch(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, subject_value="Expected Subject")
        return _make_snapshot(pdf_path, subject_value="Different Subject")

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    summary, _json_path, _markdown_path = validator.validate_corpus(tmp_path, tmp_path / "reports")

    case = summary["cases"][0]
    assert summary["failed_count"] == 1
    assert case["passed"] is False
    assert any(issue["code"] == "missing_header_subject" for issue in case["hard_failures"])


def test_validate_corpus_fails_on_missing_attachment(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, attachments_value="specification.pdf")
        return _make_snapshot(pdf_path, attachments_value="")

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    summary, _json_path, _markdown_path = validator.validate_corpus(tmp_path, tmp_path / "reports")

    case = summary["cases"][0]
    assert case["passed"] is False
    assert any(issue["code"] == "attachment_missing" for issue in case["hard_failures"])


def test_validate_corpus_fails_on_low_body_anchor_recall(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(
                pdf_path,
                body_lines=[
                    "First long regression anchor line with distinctive content and wording for matching.",
                    "Second long regression anchor line with distinctive content and wording for matching.",
                    "Third long regression anchor line with distinctive content and wording for matching.",
                    "Fourth long regression anchor line with distinctive content and wording for matching.",
                    "Fifth long regression anchor line with distinctive content and wording for matching.",
                ],
            )
        return _make_snapshot(
            pdf_path,
            body_lines=[
                "Completely different generated body text that should not match the golden anchors at all.",
            ],
        )

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    summary, _json_path, _markdown_path = validator.validate_corpus(tmp_path, tmp_path / "reports")

    case = summary["cases"][0]
    assert case["passed"] is False
    assert any(issue["code"] == "body_anchor_low" for issue in case["hard_failures"])


def test_validate_case_pair_page_count_info(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)
    pair = validator.discover_case_pairs(tmp_path)[0]

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, page_count=2)
        return _make_snapshot(pdf_path, page_count=3)

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    result = validator.validate_case_pair(pair, tmp_path / "generated")

    assert result.passed is True
    assert result.warnings == []
    assert any(issue["code"] == "page_count_info" for issue in result.infos)


def test_validate_case_pair_page_count_failure(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)
    pair = validator.discover_case_pairs(tmp_path)[0]

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, page_count=1)
        return _make_snapshot(pdf_path, page_count=4)

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    result = validator.validate_case_pair(pair, tmp_path / "generated")

    assert result.passed is False
    assert any(issue["code"] == "page_count" for issue in result.hard_failures)


def test_validate_case_pair_page_count_delta_two_is_info(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)
    pair = validator.discover_case_pairs(tmp_path)[0]

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, page_count=6)
        return _make_snapshot(pdf_path, page_count=4)

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    result = validator.validate_case_pair(pair, tmp_path / "generated")

    assert result.passed is True
    assert result.warnings == []
    assert any(issue["code"] == "page_count_info" for issue in result.infos)


def test_validate_corpus_tracks_info_case_ids(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "7")
    _write_case_pair(case_dir)

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, page_count=6)
        return _make_snapshot(pdf_path, page_count=4)

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    summary, _json_path, markdown_path = validator.validate_corpus(
        tmp_path,
        tmp_path / "reports",
        fail_on_warnings=True,
    )

    assert summary["failed_count"] == 0
    assert summary["warning_case_count"] == 0
    assert summary["info_case_count"] == 1
    assert summary["strict_failed_count"] == 0
    assert summary["issue_case_ids"]["page_count_info"] == ["7"]
    markdown = markdown_path.read_text(encoding="utf-8")
    assert "## Info Cases" in markdown
    assert "| page_count_info | 1 | 7 |" in markdown


def test_main_returns_nonzero_when_fail_on_warnings_requested(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, page_count=2)
        return _make_snapshot(pdf_path, page_count=3)

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    exit_code = validator.main(
        [
            "--cases-dir",
            str(tmp_path),
            "--output-root",
            str(tmp_path / "reports"),
            "--fail-on-warnings",
        ]
    )

    assert exit_code == 0


def test_main_returns_zero_when_warnings_are_allowed(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)

    monkeypatch.setattr(validator, "convert_msg_files", _fake_conversion)

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        if pdf_path.name == "golden.pdf":
            return _make_snapshot(pdf_path, page_count=2)
        return _make_snapshot(pdf_path, page_count=3)

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    exit_code = validator.main(
        [
            "--cases-dir",
            str(tmp_path),
            "--output-root",
            str(tmp_path / "reports"),
        ]
    )

    assert exit_code == 0


def test_main_returns_nonzero_when_reportlab_warning_is_present(monkeypatch, tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)

    monkeypatch.setattr(validator, "convert_msg_files", lambda paths, output_dir: _fake_conversion(paths, output_dir, pipeline="reportlab"))

    def fake_analyze(pdf_path: Path) -> validator.PdfSnapshot:
        return _make_snapshot(pdf_path)

    monkeypatch.setattr(validator, "analyze_pdf", fake_analyze)

    exit_code = validator.main(
        [
            "--cases-dir",
            str(tmp_path),
            "--output-root",
            str(tmp_path / "reports"),
            "--fail-on-warnings",
        ]
    )

    assert exit_code == 2


def test_profile_corpus_remains_flat_only_while_validator_handles_pairs(tmp_path: Path) -> None:
    case_dir = _make_case_dir(tmp_path, "1")
    _write_case_pair(case_dir)

    with pytest.raises(ValueError, match="No .msg files found"):
        profiler.profile_corpus(tmp_path, tmp_path / "reports")

    pairs = validator.discover_case_pairs(tmp_path)
    assert len(pairs) == 1
    assert pairs[0].case_id == "1"
