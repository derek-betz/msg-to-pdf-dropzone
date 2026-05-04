from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from fastapi.testclient import TestClient

from msg_to_pdf_dropzone import web_server
from msg_to_pdf_dropzone.feedback import (
    MAX_MESSAGE_LENGTH,
    _build_email_body,
    _coerce_bool,
    _decode_json_object,
    _env_first,
    _format_context_lines,
    _json_dumps_safe,
    _sanitize_message,
    list_feedback_entries,
    record_feedback,
    send_feedback_email,
)
from msg_to_pdf_dropzone.web_server import create_app


def test_record_feedback_persists_jsonl_and_sqlite(tmp_path: Path) -> None:
    feedback_dir = tmp_path / "feedback"

    target = record_feedback(
        {
            "source": "web",
            "category": "confusing",
            "improve": "Make the output folder clearer.",
            "helpful": "Batch conversion is quick.",
            "context": {"queuedCount": 2, "serverMode": True},
        },
        output_dir=feedback_dir,
    )

    assert target.name == "feedback.jsonl"
    assert target.exists()
    assert (feedback_dir / "feedback.sqlite3").exists()
    rows = list_feedback_entries(output_dir=feedback_dir, limit=10)
    assert len(rows) == 1
    assert rows[0]["category"] == "confusing"
    assert rows[0]["improve"] == "Make the output folder clearer."
    assert rows[0]["context"]["queuedCount"] == 2


def test_feedback_api_post_and_get(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_OUTPUT_DIR", str(tmp_path / "pdf"))
    monkeypatch.setenv("MSG_TO_PDF_FEEDBACK_READ_TOKEN", "secret")
    monkeypatch.setattr(web_server, "STAGING_DIR", tmp_path / "staging")
    monkeypatch.setattr(web_server, "send_feedback_email", lambda _payload: True)
    client = TestClient(create_app())

    posted = client.post(
        "/api/feedback",
        json={
            "category": "feature_request",
            "improve": "Add a better complete message.",
            "helpful": "The progress bars are useful.",
            "context": {"queuedCount": 1},
        },
    )

    assert posted.status_code == 200
    assert posted.json()["emailSent"] is True
    listed = client.get("/api/feedback?token=secret")
    payload = listed.json()
    assert listed.status_code == 200
    assert payload["count"] == 1
    assert payload["items"][0]["category"] == "feature_request"
    assert payload["items"][0]["helpful"] == "The progress bars are useful."
    assert payload["items"][0]["context"]["queuedCount"] == 1


def test_feedback_api_requires_content_and_read_token(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_OUTPUT_DIR", str(tmp_path / "pdf"))
    monkeypatch.delenv("MSG_TO_PDF_FEEDBACK_READ_TOKEN", raising=False)
    monkeypatch.delenv("FEEDBACK_READ_TOKEN", raising=False)
    monkeypatch.setattr(web_server, "STAGING_DIR", tmp_path / "staging")
    monkeypatch.setattr(web_server, "send_feedback_email", lambda _payload: True)
    client = TestClient(create_app())

    empty = client.post("/api/feedback", json={})
    denied = client.get("/api/feedback")

    assert empty.status_code == 400
    assert denied.status_code == 403


def test_feedback_api_reports_email_failure_without_losing_local_save(monkeypatch, tmp_path: Path) -> None:
    output_dir = tmp_path / "pdf"
    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_OUTPUT_DIR", str(output_dir))
    monkeypatch.setattr(web_server, "STAGING_DIR", tmp_path / "staging")
    monkeypatch.setattr(web_server, "send_feedback_email", lambda _payload: False)
    client = TestClient(create_app())

    response = client.post("/api/feedback", json={"message": "Save this even if email fails."})

    assert response.status_code == 200
    assert response.json()["emailSent"] is False
    rows = list_feedback_entries(output_dir=output_dir.parent / "feedback", limit=10)
    assert rows[0]["message"] == "Save this even if email fails."


def test_feedback_admin_page_requires_token_and_renders_entries(monkeypatch, tmp_path: Path) -> None:
    output_dir = tmp_path / "pdf"
    monkeypatch.setenv("MSG_TO_PDF_SERVER_MODE", "1")
    monkeypatch.setenv("MSG_TO_PDF_OUTPUT_DIR", str(output_dir))
    monkeypatch.setenv("MSG_TO_PDF_FEEDBACK_READ_TOKEN", "secret")
    monkeypatch.setattr(web_server, "STAGING_DIR", tmp_path / "staging")
    monkeypatch.setattr(web_server, "send_feedback_email", lambda _payload: True)
    client = TestClient(create_app())

    client.post("/api/feedback", json={"category": "positive", "message": "This helped."})

    denied = client.get("/feedback")
    allowed = client.get("/feedback?token=secret")

    assert denied.status_code == 403
    assert allowed.status_code == 200
    assert "msg-to-pdf-dropzone Feedback" in allowed.text
    assert "positive" in allowed.text


def test_sanitize_message_normalizes_whitespace() -> None:
    assert _sanitize_message("  hello   world  ") == "hello world"


def test_sanitize_message_truncates_at_max_length() -> None:
    long_msg = "x" * (MAX_MESSAGE_LENGTH + 100)
    result = _sanitize_message(long_msg)
    assert len(result) == MAX_MESSAGE_LENGTH


def test_sanitize_message_returns_empty_for_empty_input() -> None:
    assert _sanitize_message("") == ""


def test_coerce_bool_handles_bool_values() -> None:
    assert _coerce_bool(True) is True
    assert _coerce_bool(False) is False


def test_coerce_bool_handles_truthy_strings() -> None:
    assert _coerce_bool("true") is True
    assert _coerce_bool("1") is True
    assert _coerce_bool("yes") is True
    assert _coerce_bool("on") is True


def test_coerce_bool_handles_falsy_strings() -> None:
    assert _coerce_bool("false") is False
    assert _coerce_bool("0") is False
    assert _coerce_bool("no") is False
    assert _coerce_bool("off") is False
    assert _coerce_bool("") is False


def test_coerce_bool_handles_none() -> None:
    assert _coerce_bool(None) is False


def test_coerce_bool_handles_numeric() -> None:
    assert _coerce_bool(1) is True
    assert _coerce_bool(0) is False


def test_build_email_body_includes_source_and_category() -> None:
    payload = {
        "source": "web",
        "category": "positive",
        "improve": "Add more features.",
        "helpful": "Works great!",
        "user_agent": "Chrome/120",
        "timestamp": "2026-01-01T00:00:00+00:00",
    }
    body = _build_email_body(payload)
    assert "Source: web" in body
    assert "Category: positive" in body
    assert "Add more features." in body
    assert "Works great!" in body


def test_build_email_body_includes_optional_message() -> None:
    payload = {"message": "Extra notes here.", "source": "desktop"}
    body = _build_email_body(payload)
    assert "Extra notes here." in body
    assert "Additional notes:" in body


def test_build_email_body_includes_context_section() -> None:
    payload = {
        "source": "web",
        "context": {"queuedCount": 3, "serverMode": True},
    }
    body = _build_email_body(payload)
    assert "queuedCount" in body
    assert "App context:" in body


def test_format_context_lines_returns_empty_for_non_mapping() -> None:
    assert _format_context_lines(None) == []
    assert _format_context_lines("a string") == []
    assert _format_context_lines(42) == []


def test_format_context_lines_returns_empty_for_empty_dict() -> None:
    assert _format_context_lines({}) == []


def test_format_context_lines_formats_key_value_pairs() -> None:
    lines = _format_context_lines({"key1": "value1", "count": 5})
    assert any("key1" in line and "value1" in line for line in lines)
    assert any("count" in line for line in lines)


def test_format_context_lines_limits_to_20_entries() -> None:
    big_context = {f"k{i}": i for i in range(30)}
    lines = _format_context_lines(big_context)
    assert len(lines) == 20


def test_json_dumps_safe_serializes_standard_types() -> None:
    result = _json_dumps_safe({"ok": True, "count": 5})
    parsed = json.loads(result)
    assert parsed["ok"] is True
    assert parsed["count"] == 5


def test_json_dumps_safe_falls_back_to_str_for_non_serializable() -> None:
    result = _json_dumps_safe({"date": datetime(2026, 1, 1)})
    assert json.loads(result) is not None


def test_decode_json_object_returns_empty_for_none() -> None:
    assert _decode_json_object(None) == {}


def test_decode_json_object_returns_empty_for_empty_string() -> None:
    assert _decode_json_object("") == {}


def test_decode_json_object_returns_empty_for_invalid_json() -> None:
    assert _decode_json_object("not json") == {}


def test_decode_json_object_parses_valid_json() -> None:
    assert _decode_json_object('{"key": "value"}') == {"key": "value"}


def test_env_first_returns_first_matching_env_var(monkeypatch) -> None:
    monkeypatch.setenv("TEST_VAR_A", "hello")
    monkeypatch.delenv("TEST_VAR_B", raising=False)
    assert _env_first("TEST_VAR_B", "TEST_VAR_A") == "hello"


def test_env_first_returns_default_when_no_vars_set(monkeypatch) -> None:
    monkeypatch.delenv("TEST_MISSING_VAR", raising=False)
    assert _env_first("TEST_MISSING_VAR", default="fallback") == "fallback"


def test_env_first_returns_none_default_when_nothing_configured(monkeypatch) -> None:
    monkeypatch.delenv("TEST_MISSING_VAR", raising=False)
    assert _env_first("TEST_MISSING_VAR") is None


def test_list_feedback_entries_with_source_filter(tmp_path: Path) -> None:
    record_feedback({"source": "web", "message": "from web"}, output_dir=tmp_path)
    record_feedback({"source": "desktop", "message": "from desktop"}, output_dir=tmp_path)

    web_entries = list_feedback_entries(output_dir=tmp_path, source="web")
    assert len(web_entries) == 1
    assert web_entries[0]["source"] == "web"


def test_list_feedback_entries_with_query_filter(tmp_path: Path) -> None:
    record_feedback({"message": "unique search term xyz"}, output_dir=tmp_path)
    record_feedback({"message": "irrelevant content"}, output_dir=tmp_path)

    results = list_feedback_entries(output_dir=tmp_path, query="unique search term xyz")
    assert len(results) == 1


def test_list_feedback_entries_returns_empty_list_when_no_db(tmp_path: Path) -> None:
    results = list_feedback_entries(output_dir=tmp_path / "nonexistent")
    assert results == []


def test_list_feedback_entries_respects_limit(tmp_path: Path) -> None:
    for i in range(5):
        record_feedback({"message": f"entry {i}"}, output_dir=tmp_path)

    limited = list_feedback_entries(output_dir=tmp_path, limit=3)
    assert len(limited) == 3


def test_list_feedback_entries_respects_offset(tmp_path: Path) -> None:
    for i in range(5):
        record_feedback({"message": f"entry {i}"}, output_dir=tmp_path)

    all_entries = list_feedback_entries(output_dir=tmp_path, limit=10)
    offset_entries = list_feedback_entries(output_dir=tmp_path, limit=10, offset=3)
    assert len(offset_entries) == len(all_entries) - 3


def test_send_feedback_email_returns_false_when_disabled(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    config_path.write_text(json.dumps({"enabled": False}), encoding="utf-8")
    result = send_feedback_email({"message": "test"}, config_path=config_path)
    assert result is False


def test_send_feedback_email_returns_false_when_smtp_not_configured(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps({"enabled": True, "sender": "a@b.com", "recipients": ["c@d.com"]}),
        encoding="utf-8",
    )
    result = send_feedback_email({"message": "test"}, config_path=config_path)
    assert result is False


def test_record_feedback_normalizes_user_agent_alias(tmp_path: Path) -> None:
    target = record_feedback(
        {"source": "web", "userAgent": "Mozilla/5.0"},
        output_dir=tmp_path,
    )
    entries = list_feedback_entries(output_dir=tmp_path)
    assert entries[0]["userAgent"] == "Mozilla/5.0"
