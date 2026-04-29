from __future__ import annotations

import json
import logging
import os
import smtplib
import sqlite3
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Mapping

PACKAGE_ROOT = Path(__file__).resolve().parent
DEFAULT_FEEDBACK_DIR = PACKAGE_ROOT.parent.parent / "outputs" / "feedback"
DEFAULT_FEEDBACK_FILE = DEFAULT_FEEDBACK_DIR / "feedback.jsonl"
DEFAULT_FEEDBACK_DB = DEFAULT_FEEDBACK_DIR / "feedback.sqlite3"
DEFAULT_FEEDBACK_CONFIG_PATH = PACKAGE_ROOT.parent.parent / "data" / "reference" / "feedback" / "config.json"
SHARED_FEEDBACK_CONFIG_PATH = Path("C:/ProgramData/SharedFeedback/feedback-email.json")
MAX_MESSAGE_LENGTH = 2000
MAX_FEEDBACK_QUERY_LIMIT = 500
LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True, slots=True)
class FeedbackSMTPConfig:
    host: str
    port: int = 587
    use_tls: bool = True
    username: str | None = None
    password: str | None = None
    timeout_seconds: float = 30.0
    retries: int = 0
    backoff_factor: float = 0.0


@dataclass(frozen=True, slots=True)
class FeedbackEmailConfig:
    enabled: bool = True
    sender: str | None = None
    recipients: list[str] | None = None
    smtp: FeedbackSMTPConfig | None = None


def record_feedback(payload: Mapping[str, object], *, output_dir: Path | None = None) -> Path:
    target_dir = output_dir or DEFAULT_FEEDBACK_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    target_file = target_dir / DEFAULT_FEEDBACK_FILE.name
    entry = _normalize_feedback_entry(payload)
    with target_file.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(entry, ensure_ascii=False) + "\n")
    _record_feedback_row(entry, db_path=target_dir / DEFAULT_FEEDBACK_DB.name)
    return target_file


def list_feedback_entries(
    *,
    output_dir: Path | None = None,
    limit: int = 100,
    offset: int = 0,
    source: str | None = None,
    query: str | None = None,
) -> list[dict[str, object]]:
    target_dir = output_dir or DEFAULT_FEEDBACK_DIR
    db_path = target_dir / DEFAULT_FEEDBACK_DB.name
    if not db_path.exists():
        return []

    safe_limit = max(1, min(int(limit), MAX_FEEDBACK_QUERY_LIMIT))
    safe_offset = max(0, int(offset))
    where_parts: list[str] = []
    params: list[object] = []

    source_value = (source or "").strip()
    if source_value:
        where_parts.append("source = ?")
        params.append(source_value)

    query_value = (query or "").strip()
    if query_value:
        like_value = f"%{query_value}%"
        where_parts.append(
            "("
            "coalesce(category, '') LIKE ? OR "
            "coalesce(improve, '') LIKE ? OR "
            "coalesce(helpful, '') LIKE ? OR "
            "coalesce(message, '') LIKE ? OR "
            "coalesce(context_json, '') LIKE ?"
            ")"
        )
        params.extend([like_value, like_value, like_value, like_value, like_value])

    where_sql = f"WHERE {' AND '.join(where_parts)}" if where_parts else ""
    sql = (
        "SELECT id, timestamp_utc, source, user_agent, category, improve, helpful, message, context_json, payload_json "
        "FROM feedback_entries "
        f"{where_sql} "
        "ORDER BY id DESC "
        "LIMIT ? OFFSET ?"
    )
    params.extend([safe_limit, safe_offset])

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        _ensure_feedback_schema(conn)
        rows = conn.execute(sql, params).fetchall()

    return [
        {
            "id": int(row["id"]),
            "timestamp": row["timestamp_utc"],
            "source": row["source"],
            "userAgent": row["user_agent"],
            "category": row["category"],
            "improve": row["improve"],
            "helpful": row["helpful"],
            "message": row["message"],
            "context": _decode_json_object(row["context_json"]),
        }
        for row in rows
    ]


def send_feedback_email(payload: Mapping[str, object], *, config_path: Path | None = None) -> bool:
    cfg = load_feedback_email_config(config_path)
    if not cfg.enabled:
        LOGGER.info("Feedback email disabled; skipping send")
        return False
    if not cfg.smtp:
        LOGGER.error("Feedback SMTP configuration missing")
        return False
    if not cfg.sender:
        LOGGER.error("Feedback sender email is not configured")
        return False
    if not cfg.recipients:
        LOGGER.error("Feedback recipients are not configured")
        return False

    message = EmailMessage()
    message["Subject"] = "MSG to PDF Dropzone Feedback"
    message["From"] = cfg.sender
    message["To"] = ", ".join(cfg.recipients)
    message.set_content(_build_email_body(payload))

    try:
        _send_smtp_message(message, cfg.smtp)
        return True
    except Exception as exc:  # pragma: no cover - defensive logging
        LOGGER.error("Feedback email failed: %s", exc)
        return False


def load_feedback_email_config(path: Path | None = None) -> FeedbackEmailConfig:
    config_path = path or _feedback_config_path()
    raw: dict[str, Any] = {}
    if config_path.exists():
        with config_path.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)

    smtp_section = raw.get("smtp") or {}
    smtp_host = smtp_section.get("host") or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_HOST", "FEEDBACK_SMTP_HOST", default="")
    smtp_cfg = None
    if smtp_host or any(key in os.environ for key in ("MSG_TO_PDF_FEEDBACK_SMTP_HOST", "FEEDBACK_SMTP_HOST")):
        smtp_cfg = _load_smtp_config(smtp_section)

    return FeedbackEmailConfig(
        enabled=_coerce_bool(
            raw.get(
                "enabled",
                _env_first("MSG_TO_PDF_FEEDBACK_EMAIL_ENABLED", "FEEDBACK_EMAIL_ENABLED", default="true"),
            )
        ),
        sender=raw.get("sender")
        or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_SENDER", "FEEDBACK_SMTP_SENDER", default=None),
        recipients=list(raw.get("recipients") or _env_recipients()),
        smtp=smtp_cfg,
    )


def _normalize_feedback_entry(payload: Mapping[str, object]) -> dict[str, Any]:
    entry: dict[str, Any] = dict(payload)
    entry.setdefault("timestamp", datetime.now(timezone.utc).isoformat())
    if not entry.get("user_agent") and entry.get("userAgent"):
        entry["user_agent"] = str(entry.get("userAgent") or "").strip()

    for key in ("message", "improve", "helpful", "category"):
        value = entry.get(key)
        if isinstance(value, str):
            entry[key] = _sanitize_message(value)
    for key in ("source", "user_agent"):
        value = entry.get(key)
        if value is not None and not isinstance(value, str):
            entry[key] = str(value)
    context = entry.get("context")
    if context is None:
        entry["context"] = {}
    elif isinstance(context, Mapping):
        entry["context"] = dict(context)
    else:
        entry["context"] = {"value": str(context)}
    return entry


def _ensure_feedback_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS feedback_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp_utc TEXT NOT NULL,
            source TEXT NOT NULL,
            user_agent TEXT,
            category TEXT,
            improve TEXT,
            helpful TEXT,
            message TEXT,
            context_json TEXT,
            payload_json TEXT NOT NULL
        )
        """
    )
    existing_columns = {
        row[1]
        for row in conn.execute("PRAGMA table_info(feedback_entries)").fetchall()
    }
    if "category" not in existing_columns:
        conn.execute("ALTER TABLE feedback_entries ADD COLUMN category TEXT")
    if "context_json" not in existing_columns:
        conn.execute("ALTER TABLE feedback_entries ADD COLUMN context_json TEXT")
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_feedback_entries_timestamp ON feedback_entries(timestamp_utc DESC)"
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_feedback_entries_source ON feedback_entries(source)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_feedback_entries_category ON feedback_entries(category)")


def _record_feedback_row(entry: Mapping[str, Any], *, db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    timestamp = str(entry.get("timestamp") or datetime.now(timezone.utc).isoformat())
    source = str(entry.get("source") or "unknown")
    user_agent = str(entry.get("user_agent") or "").strip() or None
    category = str(entry.get("category") or "").strip() or None
    improve = str(entry.get("improve") or "").strip() or None
    helpful = str(entry.get("helpful") or "").strip() or None
    message = str(entry.get("message") or "").strip() or None
    context_json = _json_dumps_safe(entry.get("context") or {})
    payload_json = json.dumps(dict(entry), ensure_ascii=False)

    with sqlite3.connect(db_path) as conn:
        _ensure_feedback_schema(conn)
        conn.execute(
            """
            INSERT INTO feedback_entries (
                timestamp_utc,
                source,
                user_agent,
                category,
                improve,
                helpful,
                message,
                context_json,
                payload_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (timestamp, source, user_agent, category, improve, helpful, message, context_json, payload_json),
        )
        conn.commit()


def _load_smtp_config(raw: Mapping[str, Any]) -> FeedbackSMTPConfig:
    data = dict(raw)
    return FeedbackSMTPConfig(
        host=str(data.get("host") or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_HOST", "FEEDBACK_SMTP_HOST", default="")),
        port=int(data.get("port") or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_PORT", "FEEDBACK_SMTP_PORT", default="587")),
        use_tls=_coerce_bool(
            data.get(
                "use_tls",
                _env_first("MSG_TO_PDF_FEEDBACK_SMTP_USE_TLS", "FEEDBACK_SMTP_USE_TLS", default="true"),
            )
        ),
        username=data.get("username")
        or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_USERNAME", "FEEDBACK_SMTP_USERNAME", default=None),
        password=data.get("password")
        or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_PASSWORD", "FEEDBACK_SMTP_PASSWORD", default=None),
        timeout_seconds=float(
            data.get("timeout_seconds")
            or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_TIMEOUT_SECONDS", "FEEDBACK_SMTP_TIMEOUT_SECONDS", default="30")
        ),
        retries=int(
            data.get("retries")
            or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_RETRIES", "FEEDBACK_SMTP_RETRIES", default="0")
        ),
        backoff_factor=float(
            data.get("backoff_factor")
            or _env_first("MSG_TO_PDF_FEEDBACK_SMTP_BACKOFF_FACTOR", "FEEDBACK_SMTP_BACKOFF_FACTOR", default="0")
        ),
    )


def _send_smtp_message(message: EmailMessage, smtp_cfg: FeedbackSMTPConfig) -> None:
    attempts = max(0, smtp_cfg.retries) + 1
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            with smtplib.SMTP(smtp_cfg.host, smtp_cfg.port, timeout=smtp_cfg.timeout_seconds) as smtp:
                if smtp_cfg.use_tls:
                    smtp.starttls()
                if smtp_cfg.username and smtp_cfg.password:
                    smtp.login(smtp_cfg.username, smtp_cfg.password)
                smtp.send_message(message)
            return
        except Exception as exc:
            last_error = exc
            if attempt >= attempts - 1:
                break
            if smtp_cfg.backoff_factor > 0:
                time.sleep(smtp_cfg.backoff_factor * (2**attempt))
    assert last_error is not None
    raise last_error


def _build_email_body(payload: Mapping[str, object]) -> str:
    user_agent = str(payload.get("user_agent") or payload.get("userAgent") or "").strip()
    timestamp = str(payload.get("timestamp") or "").strip()
    category = _sanitize_message(str(payload.get("category") or "other"))
    improve = _sanitize_message(str(payload.get("improve") or ""))
    helpful = _sanitize_message(str(payload.get("helpful") or ""))
    message = _sanitize_message(str(payload.get("message") or ""))
    context_lines = _format_context_lines(payload.get("context"))

    lines = [
        "New feedback received for MSG to PDF Dropzone.",
        "",
        f"Source: {payload.get('source') or 'web'}",
        f"Category: {category or 'n/a'}",
        f"User Agent: {user_agent or 'n/a'}",
        f"Submitted (UTC): {timestamp or 'n/a'}",
        "",
        "What could be better about this application?",
        improve or "(no response)",
        "",
        "What is good or helpful about this application?",
        helpful or "(no response)",
        "",
    ]
    if message:
        lines.extend(["Additional notes:", message, ""])
    if context_lines:
        lines.extend(["App context:", *context_lines, ""])
    return "\n".join(lines).strip() + "\n"


def _json_dumps_safe(value: object) -> str:
    try:
        return json.dumps(value, ensure_ascii=False, default=str)
    except TypeError:
        return json.dumps({"value": str(value)}, ensure_ascii=False)


def _decode_json_object(value: object) -> object:
    if not value:
        return {}
    try:
        return json.loads(str(value))
    except json.JSONDecodeError:
        return {}


def _format_context_lines(context: object) -> list[str]:
    if not isinstance(context, Mapping):
        return []
    lines: list[str] = []
    for key, value in list(context.items())[:20]:
        label = _sanitize_message(str(key))
        if isinstance(value, (dict, list, tuple)):
            formatted = _json_dumps_safe(value)
        else:
            formatted = str(value)
        formatted = _sanitize_message(formatted)
        lines.append(f"- {label}: {formatted or 'n/a'}")
    return lines


def _feedback_config_path() -> Path:
    configured = _env_first("MSG_TO_PDF_FEEDBACK_CONFIG_PATH", "FEEDBACK_CONFIG_PATH", default="")
    if configured:
        return Path(configured).expanduser()
    if SHARED_FEEDBACK_CONFIG_PATH.exists():
        return SHARED_FEEDBACK_CONFIG_PATH
    return DEFAULT_FEEDBACK_CONFIG_PATH


def _env_recipients() -> list[str]:
    raw = _env_first("MSG_TO_PDF_FEEDBACK_SMTP_RECIPIENTS", "FEEDBACK_SMTP_RECIPIENTS", default="") or ""
    return [email.strip() for email in raw.split(",") if email.strip()]


def _env_first(*names: str, default: str | None = None) -> str | None:
    for name in names:
        value = os.environ.get(name)
        if value is not None:
            return value
    return default


def _sanitize_message(message: str) -> str:
    cleaned = " ".join(message.strip().split())
    if len(cleaned) > MAX_MESSAGE_LENGTH:
        return cleaned[:MAX_MESSAGE_LENGTH].rstrip()
    return cleaned


def _coerce_bool(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip().lower() not in {"0", "false", "no", "off", ""}
    return bool(value)
