from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path

from msg_to_pdf_dropzone.models import EmailRecord
from msg_to_pdf_dropzone.thread_logic import (
    build_pdf_filename,
    get_latest_thread_dates,
    make_unique_path,
    normalize_thread_subject,
    sanitize_filename_part,
)


def _record(subject: str, sent_at: datetime) -> EmailRecord:
    return EmailRecord(
        source_path=Path("sample.msg"),
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


def test_normalize_thread_subject_removes_reply_prefixes() -> None:
    assert normalize_thread_subject("RE: FW: Fwd: Quarterly Update") == "quarterly update"


def test_get_latest_thread_dates_uses_latest_date_per_thread() -> None:
    records = [
        _record("Re: Weekly Sync", datetime(2026, 1, 10, tzinfo=timezone.utc)),
        _record("Weekly Sync", datetime(2026, 1, 15, tzinfo=timezone.utc)),
        _record("Budget", datetime(2026, 1, 12, tzinfo=timezone.utc)),
    ]

    latest = get_latest_thread_dates(records, local_tz=timezone.utc)

    assert latest["weekly sync"].isoformat() == "2026-01-15"
    assert latest["budget"].isoformat() == "2026-01-12"


def test_get_latest_thread_dates_uses_local_calendar_day_when_requested() -> None:
    records = [
        _record("Weekly Sync", datetime(2026, 1, 16, 1, 0, tzinfo=timezone.utc)),
    ]

    latest = get_latest_thread_dates(records, local_tz=timezone(timedelta(hours=-5)))

    assert latest["weekly sync"].isoformat() == "2026-01-15"


def test_build_pdf_filename_prefixes_date_and_sanitizes_subject() -> None:
    file_name = build_pdf_filename("Q1: Planning / Kickoff?", datetime(2026, 2, 20).date())
    assert file_name == "2026-02-20_Q1_ Planning _ Kickoff_.pdf"


def test_sanitize_filename_part_defaults_when_empty() -> None:
    assert sanitize_filename_part("   ") == "No Subject"


def test_sanitize_filename_part_preserves_internal_double_spaces() -> None:
    assert sanitize_filename_part("DES  2101166  SR 127") == "DES  2101166  SR 127"


def test_sanitize_filename_part_truncates_at_max_length() -> None:
    long_name = "A" * 150
    result = sanitize_filename_part(long_name)
    assert len(result) <= 120


def test_sanitize_filename_part_replaces_invalid_chars() -> None:
    assert sanitize_filename_part("file<name>:test") == "file_name__test"


def test_normalize_thread_subject_returns_no_subject_for_empty_string() -> None:
    assert normalize_thread_subject("") == "no-subject"
    assert normalize_thread_subject("   ") == "no-subject"


def test_make_unique_path_returns_unchanged_when_no_conflict(tmp_path: Path) -> None:
    path = tmp_path / "output.pdf"
    assert make_unique_path(path) == path


def test_make_unique_path_adds_counter_suffix_on_conflict(tmp_path: Path) -> None:
    existing = tmp_path / "output.pdf"
    existing.write_text("x", encoding="utf-8")
    new_path = make_unique_path(existing)
    assert new_path == tmp_path / "output (2).pdf"


def test_make_unique_path_increments_counter_past_existing_duplicates(tmp_path: Path) -> None:
    for name in ("output.pdf", "output (2).pdf"):
        (tmp_path / name).write_text("x", encoding="utf-8")
    result = make_unique_path(tmp_path / "output.pdf")
    assert result == tmp_path / "output (3).pdf"
