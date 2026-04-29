from __future__ import annotations

from pathlib import Path

from fastapi.testclient import TestClient

from msg_to_pdf_dropzone import web_server
from msg_to_pdf_dropzone.feedback import list_feedback_entries, record_feedback
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
