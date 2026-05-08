from __future__ import annotations

import argparse
import asyncio
import atexit
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import datetime, timezone
import hmac
import html
from importlib.metadata import PackageNotFoundError, version
import json
import os
from pathlib import Path
import queue
import subprocess
import tempfile
import threading
from typing import Any
from uuid import uuid4
import webbrowser

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn

from .converter import MAX_FILES_PER_BATCH, ConversionResult, convert_msg_files
from .feedback import list_feedback_entries, record_feedback, send_feedback_email
from .msg_parser import parse_msg_file
from .outlook_selection import extract_selected_outlook_messages
from .task_events import TaskEvent, TaskMetaValue, build_batch_meta_for_paths, default_task_id_for_path, emit_task_event
from .thread_logic import DEFAULT_FILENAME_STYLE, build_pdf_filename, get_latest_thread_dates, normalize_filename_style

PACKAGE_ROOT = Path(__file__).resolve().parent
WEB_UI_DIR = PACKAGE_ROOT / "web_ui"
ASSET_DIR = PACKAGE_ROOT / "assets"
APP_NAME = "msg-to-pdf-dropzone"
RELEASE_INFO_FILE = "_release.json"


def _package_version() -> str:
    try:
        return version(APP_NAME)
    except PackageNotFoundError:
        return "0.1.0"


def _read_release_info() -> dict[str, object]:
    release_path = PACKAGE_ROOT / RELEASE_INFO_FILE
    try:
        raw_payload = json.loads(release_path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}
    if not isinstance(raw_payload, dict):
        return {}
    return raw_payload


def _git_revision() -> str | None:
    env_revision = (
        os.getenv("MSG_TO_PDF_APP_REVISION")
        or os.getenv("MSG_TO_PDF_SOURCE_REVISION")
        or ""
    ).strip()
    if env_revision:
        return env_revision

    release_info = _read_release_info()
    for key in ("sourceRevision", "commit", "gitCommit"):
        value = release_info.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()

    repo_root = PACKAGE_ROOT.parents[1]
    try:
        result = subprocess.run(
            ["git", "-C", str(repo_root), "rev-parse", "HEAD"],
            check=True,
            capture_output=True,
            text=True,
            timeout=5,
        )
    except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return None
    revision = result.stdout.strip()
    return revision or None


def build_version_payload() -> dict[str, object]:
    release_info = _read_release_info()
    payload: dict[str, object] = {
        "appName": APP_NAME,
        "appVersion": _package_version(),
        "sourceRevision": _git_revision(),
    }
    for key in ("sourceBranch", "deployedAt", "deployedBy"):
        value = release_info.get(key)
        if isinstance(value, str) and value.strip():
            payload[key] = value.strip()
    return payload


def _env_flag(name: str, default: bool = False) -> bool:
    raw_value = os.getenv(name, "").strip().lower()
    if not raw_value:
        return default
    return raw_value in {"1", "true", "yes", "on"}


def _env_path(name: str) -> Path | None:
    raw_value = os.getenv(name, "").strip()
    if not raw_value:
        return None
    return Path(raw_value).expanduser()


def _env_csv(name: str) -> tuple[str, ...]:
    raw_value = os.getenv(name, "").strip()
    if not raw_value:
        return ()
    return tuple(part.strip() for part in raw_value.split(",") if part.strip())


@dataclass(frozen=True, slots=True)
class HostingSettings:
    server_mode: bool
    default_output_dir: Path | None
    disable_outlook_import: bool
    disable_output_picker: bool


@dataclass(frozen=True, slots=True)
class ServerSettings:
    host: str
    port: int
    tls_certfile: Path | None
    tls_keyfile: Path | None
    tls_server_names: tuple[str, ...]

    @property
    def tls_enabled(self) -> bool:
        return self.tls_certfile is not None and self.tls_keyfile is not None


def load_hosting_settings() -> HostingSettings:
    server_mode = _env_flag("MSG_TO_PDF_SERVER_MODE", False)
    default_output_dir = _env_path("MSG_TO_PDF_OUTPUT_DIR")
    disable_outlook_import = _env_flag("MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT", server_mode)
    disable_output_picker = _env_flag(
        "MSG_TO_PDF_DISABLE_OUTPUT_PICKER",
        server_mode or default_output_dir is not None,
    )
    return HostingSettings(
        server_mode=server_mode,
        default_output_dir=default_output_dir,
        disable_outlook_import=disable_outlook_import,
        disable_output_picker=disable_output_picker,
    )


STAGING_DIR = _env_path("MSG_TO_PDF_STAGING_DIR") or (Path(tempfile.gettempdir()) / "msg-to-pdf-browser-staging")


def _is_loopback_host(host: str) -> bool:
    normalized = (host or "").strip().lower()
    return normalized in {"127.0.0.1", "localhost"}


def load_server_settings(*, host: str | None = None, port: int | None = None) -> ServerSettings:
    resolved_host = (host or os.getenv("MSG_TO_PDF_HOST", "127.0.0.1")).strip() or "127.0.0.1"
    resolved_port = port or int(os.getenv("MSG_TO_PDF_PORT", "8765"))
    return ServerSettings(
        host=resolved_host,
        port=resolved_port,
        tls_certfile=_env_path("MSG_TO_PDF_TLS_CERTFILE"),
        tls_keyfile=_env_path("MSG_TO_PDF_TLS_KEYFILE"),
        tls_server_names=_env_csv("MSG_TO_PDF_SERVER_NAMES"),
    )


def validate_startup_contract(settings: ServerSettings) -> None:
    if bool(settings.tls_certfile) != bool(settings.tls_keyfile):
        raise SystemExit("msg-to-pdf-dropzone TLS startup requires both MSG_TO_PDF_TLS_CERTFILE and MSG_TO_PDF_TLS_KEYFILE.")
    if settings.tls_enabled:
        assert settings.tls_certfile is not None
        assert settings.tls_keyfile is not None
        if not settings.tls_certfile.exists():
            raise SystemExit(f"msg-to-pdf-dropzone TLS cert file was not found: {settings.tls_certfile}")
        if not settings.tls_keyfile.exists():
            raise SystemExit(f"msg-to-pdf-dropzone TLS key file was not found: {settings.tls_keyfile}")
        return
    app_env = (os.getenv("APP_ENV", "production").strip() or "production").lower()
    if app_env == "production" and not _is_loopback_host(settings.host):
        raise SystemExit(
            "msg-to-pdf-dropzone production startup requires loopback binding unless direct TLS is configured."
        )


@dataclass(slots=True)
class StagedFile:
    id: str
    task_id: str
    path: Path
    name: str
    size_bytes: int
    source: str
    created_at: str
    stage: str = "files_accepted"
    pipeline: str | None = None
    error: str | None = None
    success: bool | None = None
    output_name: str | None = None
    output_path: str | None = None
    cleanup: bool = True

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {
            "id": self.id,
            "taskId": self.task_id,
            "name": self.name,
            "sizeBytes": self.size_bytes,
            "source": self.source,
            "createdAt": self.created_at,
            "stage": self.stage,
        }
        if self.pipeline is not None:
            payload["pipeline"] = self.pipeline
        if self.error is not None:
            payload["error"] = self.error
        if self.success is not None:
            payload["success"] = self.success
        if self.output_name is not None:
            payload["outputName"] = self.output_name
        if self.output_path is not None:
            payload["outputPath"] = self.output_path
        return payload


class EventBroker:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._subscribers: set[queue.Queue[str]] = set()

    def publish_task_event(self, event: TaskEvent) -> None:
        self._broadcast(json.dumps(event.to_dict(), ensure_ascii=True))

    def publish_status(self, *, kind: str, message: str) -> None:
        self._broadcast(json.dumps({"kind": kind, "message": message}, ensure_ascii=True))

    def _broadcast(self, payload: str) -> None:
        with self._lock:
            subscribers = list(self._subscribers)
        for subscriber in subscribers:
            try:
                subscriber.put_nowait(payload)
            except queue.Full:
                continue

    def subscribe(self) -> queue.Queue[str]:
        subscriber: queue.Queue[str] = queue.Queue()
        with self._lock:
            self._subscribers.add(subscriber)
        return subscriber

    def unsubscribe(self, subscriber: queue.Queue[str]) -> None:
        with self._lock:
            self._subscribers.discard(subscriber)


class StageStore:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._items: dict[str, StagedFile] = {}
        self._staging_dir = STAGING_DIR
        self._staging_dir.mkdir(parents=True, exist_ok=True)
        atexit.register(self._cleanup_on_exit)

    def snapshot(self) -> list[dict[str, object]]:
        with self._lock:
            return [item.to_dict() for item in self._items.values()]

    def count(self) -> int:
        with self._lock:
            return len(self._items)

    def active_count(self) -> int:
        with self._lock:
            return sum(1 for item in self._items.values() if item.stage != "complete")

    def remaining_slots(self) -> int:
        return max(0, MAX_FILES_PER_BATCH - self.active_count())

    async def stage_uploads(self, uploads: list[UploadFile], *, source: str = "upload") -> list[StagedFile]:
        staged: list[StagedFile] = []
        normalized_source = "outlook" if source == "outlook" else "upload"
        for upload in uploads[: self.remaining_slots()]:
            try:
                if not upload.filename or Path(upload.filename).suffix.lower() != ".msg":
                    continue
                self._staging_dir.mkdir(parents=True, exist_ok=True)
                target = self._staging_dir / f"{uuid4().hex}-{Path(upload.filename).name}"
                with target.open("wb") as handle:
                    while True:
                        chunk = await upload.read(1024 * 1024)
                        if not chunk:
                            break
                        handle.write(chunk)
                staged.append(self._register(target, name=Path(upload.filename).name, source=normalized_source))
            finally:
                await upload.close()
        return staged

    def stage_paths(self, paths: Iterable[Path], *, source: str) -> list[StagedFile]:
        staged: list[StagedFile] = []
        for path in list(paths)[: self.remaining_slots()]:
            if path.suffix.lower() != ".msg" or not path.exists():
                continue
            staged.append(self._register(path, name=path.name, source=source))
        return staged

    def remove(self, ids: Iterable[str]) -> None:
        with self._lock:
            items = [self._items.pop(item_id) for item_id in ids if item_id in self._items]
        for item in items:
            self._delete_staged_file(item)

    def clear(self) -> None:
        with self._lock:
            items = list(self._items.values())
            self._items.clear()
        for item in items:
            self._delete_staged_file(item)

    def resolve_items(self, ids: list[str]) -> list[StagedFile]:
        with self._lock:
            items = [self._items[item_id] for item_id in ids if item_id in self._items]
        if not items:
            raise HTTPException(status_code=400, detail="No queued .msg files were selected.")
        return items

    def resolve_convertible_items(self, ids: list[str]) -> list[StagedFile]:
        items = self.resolve_items(ids)
        convertible = [item for item in items if item.stage != "complete"]
        if not convertible:
            raise HTTPException(status_code=400, detail="Selected .msg files are already complete.")
        return convertible

    def apply_task_event(self, event: TaskEvent) -> None:
        output_name = None
        output_path = None
        if event.meta:
            raw_output_name = event.meta.get("outputName")
            if isinstance(raw_output_name, str) and raw_output_name:
                output_name = raw_output_name
            raw_output_path = event.meta.get("outputPath")
            if isinstance(raw_output_path, str) and raw_output_path:
                output_path = raw_output_path

        with self._lock:
            item = next((candidate for candidate in self._items.values() if candidate.task_id == event.task_id), None)
            if item is None:
                return
            item.stage = event.stage
            item.pipeline = event.pipeline
            item.error = event.error
            item.success = event.success
            if output_name is not None:
                item.output_name = output_name
            if output_path is not None:
                item.output_path = output_path

    def refresh_output_previews(self, *, filename_style: str | None = DEFAULT_FILENAME_STYLE) -> None:
        resolved_style = normalize_filename_style(filename_style)
        with self._lock:
            items = [item for item in self._items.values() if item.stage != "complete"]

        records = []
        records_by_path: dict[Path, Any] = {}
        for item in items:
            try:
                record = parse_msg_file(item.path)
            except Exception:
                continue
            records.append(record)
            records_by_path[item.path.resolve()] = record

        if not records:
            return

        latest_thread_dates = get_latest_thread_dates(records)
        with self._lock:
            for item in self._items.values():
                record = records_by_path.get(item.path.resolve())
                if record is None:
                    continue
                latest_date = latest_thread_dates.get(record.thread_key)
                if latest_date is None:
                    continue
                item.output_name = build_pdf_filename(
                    record.subject,
                    latest_date,
                    filename_style=resolved_style,
                    sender=record.sender,
                )

    def _register(self, path: Path, *, name: str, source: str) -> StagedFile:
        item = StagedFile(
            id=uuid4().hex[:12],
            task_id=default_task_id_for_path(path),
            path=path.resolve(),
            name=name,
            size_bytes=path.stat().st_size if path.exists() else 0,
            source=source,
            created_at=datetime.now().astimezone().isoformat(timespec="seconds"),
        )
        with self._lock:
            if len(self._items) >= MAX_FILES_PER_BATCH:
                raise HTTPException(status_code=400, detail=f"Only {MAX_FILES_PER_BATCH} files can be queued.")
            self._items[item.id] = item
        return item

    def _delete_staged_file(self, item: StagedFile) -> None:
        if not item.cleanup:
            return
        try:
            if item.path.exists():
                item.path.unlink()
        except Exception:
            pass

    def _cleanup_on_exit(self) -> None:
        self.clear()


class RemoveRequest(BaseModel):
    ids: list[str]


class ConvertRequest(BaseModel):
    ids: list[str]
    output_dir: str | None = None
    filename_style: str = DEFAULT_FILENAME_STYLE


class FilenameStyleRequest(BaseModel):
    filename_style: str = DEFAULT_FILENAME_STYLE


class OpenOutputFolderRequest(BaseModel):
    output_dir: str | None = None


class RevealOutputFileRequest(BaseModel):
    output_path: str | None = None


class PreviewRequest(BaseModel):
    pipeline: str = "outlook_edge"


class FeedbackRequest(BaseModel):
    category: str = "other"
    improve: str = ""
    helpful: str = ""
    message: str = ""
    context: dict[str, Any] | None = None


def choose_output_directory() -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askdirectory(title="Choose where to save converted PDFs")
    finally:
        root.destroy()
    return selected or None


def open_output_directory(path_value: str) -> bool:
    target = Path(path_value).expanduser()
    if not target.exists() or not target.is_dir():
        return False

    startfile = getattr(os, "startfile", None)
    if callable(startfile):
        startfile(str(target))
        return True

    try:
        subprocess.Popen(["xdg-open", str(target)])
    except OSError:
        return False
    return True


def reveal_output_file(path_value: str) -> bool:
    target = Path(path_value).expanduser()
    if target.suffix.lower() != ".pdf" or not target.exists() or not target.is_file():
        return False

    if os.name == "nt":
        try:
            subprocess.Popen(["explorer.exe", f"/select,{target}"])
        except OSError:
            return False
        return True

    return open_output_directory(str(target.parent))


def build_batch_meta(items: list[StagedFile]) -> dict[Path, dict[str, TaskMetaValue]]:
    return build_batch_meta_for_paths([item.path for item in items])


def _resolve_output_dir(request_output_dir: str | None, settings: HostingSettings) -> Path:
    raw_value = (request_output_dir or "").strip()
    if raw_value:
        output_dir = Path(raw_value).expanduser()
    elif settings.default_output_dir is not None:
        output_dir = settings.default_output_dir
    else:
        raise HTTPException(status_code=400, detail="Choose an output folder before converting.")
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def summarize_result(result: ConversionResult) -> dict[str, object]:
    return {
        "requestedCount": result.requested_count,
        "convertedFiles": [str(path) for path in result.converted_files],
        "errors": list(result.errors),
        "skippedFiles": [str(path) for path in result.skipped_files],
        "totalSeconds": result.total_seconds,
        "timingLines": list(result.timing_lines),
    }


def _feedback_output_dir(hosting: HostingSettings) -> Path:
    if hosting.default_output_dir is not None:
        return hosting.default_output_dir.parent / "feedback"
    return Path(os.getenv("MSG_TO_PDF_FEEDBACK_DIR", "") or (PACKAGE_ROOT.parent.parent / "outputs" / "feedback")).expanduser()


def _feedback_read_allowed(token: str | None) -> bool:
    expected_token = (
        os.getenv("MSG_TO_PDF_FEEDBACK_READ_TOKEN")
        or os.getenv("FEEDBACK_READ_TOKEN")
        or ""
    ).strip()
    provided_token = (token or "").strip()
    if not expected_token or not provided_token:
        return False
    return hmac.compare_digest(provided_token, expected_token)


def _feedback_request_context(
    request: Request,
    payload: FeedbackRequest,
    *,
    hosting: HostingSettings,
    stage_store: StageStore,
) -> dict[str, Any]:
    default_output_dir = hosting.default_output_dir
    context: dict[str, Any] = {
        "appName": "msg-to-pdf-dropzone",
        "url": str(request.url),
        "path": request.url.path,
        "referrer": request.headers.get("Referer", ""),
        "host": request.headers.get("Host", ""),
        "serverMode": hosting.server_mode,
        "defaultOutputDir": str(default_output_dir) if default_output_dir else None,
        "queuedCount": stage_store.count(),
        "activeQueueCount": stage_store.active_count(),
        "disableOutlookImport": hosting.disable_outlook_import,
        "disableOutputPicker": hosting.disable_output_picker,
    }
    if payload.context:
        context.update(payload.context)
    return context


def _render_feedback_admin_page(items: list[dict[str, object]]) -> str:
    rows = []
    for item in items:
        context_text = _json_dumps_for_html(item.get("context") or {})
        rows.append(
            "<tr>"
            f"<td>{html.escape(str(item.get('timestamp') or ''))}</td>"
            f"<td>{html.escape(str(item.get('category') or ''))}</td>"
            f"<td>{html.escape(str(item.get('improve') or ''))}</td>"
            f"<td>{html.escape(str(item.get('helpful') or ''))}</td>"
            f"<td>{html.escape(str(item.get('message') or ''))}</td>"
            f"<td><pre>{html.escape(context_text)}</pre></td>"
            "</tr>"
        )
    body = "\n".join(rows) or '<tr><td colspan="6">No feedback entries found.</td></tr>'
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>msg-to-pdf-dropzone Feedback</title>
  <style>
    body {{ font-family: Segoe UI, sans-serif; margin: 24px; color: #172335; background: #f5efe0; }}
    table {{ border-collapse: collapse; width: 100%; background: #fffaf0; }}
    th, td {{ border: 1px solid #d8cdb7; padding: 8px; text-align: left; vertical-align: top; }}
    th {{ background: #e9dcc4; }}
    pre {{ margin: 0; white-space: pre-wrap; max-width: 420px; }}
  </style>
</head>
<body>
  <h1>msg-to-pdf-dropzone Feedback</h1>
  <p>Showing latest {len(items)} saved feedback entr{'y' if len(items) == 1 else 'ies'}.</p>
  <table>
    <thead>
      <tr><th>Submitted</th><th>Category</th><th>Could Be Better</th><th>Helpful</th><th>Notes</th><th>Context</th></tr>
    </thead>
    <tbody>{body}</tbody>
  </table>
</body>
</html>"""


def _json_dumps_for_html(value: object) -> str:
    try:
        return json.dumps(value, indent=2, ensure_ascii=False, default=str)
    except TypeError:
        return str(value)


async def publish_preview_sequence(event_broker: EventBroker, *, pipeline: str) -> None:
    task_id = f"preview-{uuid4().hex[:10]}"
    file_name = "Quarterly Planning.msg"
    batch_meta: dict[str, TaskMetaValue] = {
        "batchId": "preview-batch",
        "batchIndex": 1,
        "batchSize": 1,
        "outputDirLabel": "Preview",
        "outputName": "2026-04-04_Quarterly Planning.pdf",
    }
    sequence = [
        ("drop_received", None, 0.35),
        ("files_accepted", None, 0.45),
        ("output_folder_selected", None, 0.45),
        ("parse_started", None, 0.6),
        ("filename_built", None, 0.55),
        ("pdf_pipeline_started", None, 0.55),
        ("pipeline_selected", pipeline, 0.7),
        ("pdf_written", pipeline, 0.7),
        ("deliver_started", pipeline, 0.55),
        ("complete", pipeline, 0.0),
    ]
    event_broker.publish_status(kind="status", message="Running the mailroom preview sequence.")
    for stage, stage_pipeline, pause_seconds in sequence:
        emit_task_event(
            event_broker.publish_task_event,
            task_id=task_id,
            stage=stage,
            file_name=file_name,
            pipeline=stage_pipeline,
            success=True if stage == "complete" else None,
            meta=batch_meta,
        )
        if pause_seconds > 0:
            await asyncio.sleep(pause_seconds)


def create_app() -> FastAPI:
    app = FastAPI(title="MSG to PDF Browser")
    app.state.stage_store = StageStore()
    app.state.event_broker = EventBroker()
    app.state.hosting_settings = load_hosting_settings()

    def publish_task_event(event: TaskEvent) -> None:
        app.state.stage_store.apply_task_event(event)
        app.state.event_broker.publish_task_event(event)

    app.mount("/static", StaticFiles(directory=str(WEB_UI_DIR)), name="static")
    app.mount("/assets", StaticFiles(directory=str(ASSET_DIR)), name="assets")

    @app.get("/")
    async def index() -> FileResponse:
        return FileResponse(WEB_UI_DIR / "index.html")

    @app.get("/api/health")
    async def health() -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        return {
            "ok": True,
            "queuedCount": stage_store.count(),
            "maxFiles": MAX_FILES_PER_BATCH,
            **build_version_payload(),
        }

    @app.get("/api/version")
    async def app_version() -> dict[str, object]:
        return build_version_payload()

    @app.get("/api/settings")
    async def settings() -> dict[str, object]:
        hosting: HostingSettings = app.state.hosting_settings
        default_output_dir = hosting.default_output_dir
        return {
            "maxFiles": MAX_FILES_PER_BATCH,
            "serverMode": hosting.server_mode,
            "defaultOutputDir": str(default_output_dir) if default_output_dir else None,
            "defaultOutputDirLabel": (
                default_output_dir.name or str(default_output_dir) if default_output_dir else None
            ),
            "capabilities": {
                "nativeOutputPicker": not hosting.disable_output_picker,
                "outlookImport": not hosting.disable_outlook_import,
            },
        }

    @app.get("/api/feedback")
    async def get_feedback(
        limit: int = 100,
        offset: int = 0,
        source: str | None = None,
        q: str | None = None,
        token: str | None = None,
    ) -> dict[str, object]:
        if not _feedback_read_allowed(token):
            raise HTTPException(status_code=403, detail="Feedback database access is restricted.")
        items = list_feedback_entries(
            output_dir=_feedback_output_dir(app.state.hosting_settings),
            limit=limit,
            offset=offset,
            source=source,
            query=q,
        )
        return {"items": items, "count": len(items), "limit": limit, "offset": offset}

    @app.get("/feedback")
    @app.get("/feedback/admin")
    async def feedback_admin(
        limit: int = 100,
        offset: int = 0,
        token: str | None = None,
    ) -> HTMLResponse:
        if not _feedback_read_allowed(token):
            return HTMLResponse("Feedback database access is restricted.", status_code=403)
        items = list_feedback_entries(
            output_dir=_feedback_output_dir(app.state.hosting_settings),
            limit=limit,
            offset=offset,
        )
        return HTMLResponse(_render_feedback_admin_page(items))

    @app.post("/api/feedback")
    async def submit_feedback(request: Request, feedback: FeedbackRequest) -> dict[str, object]:
        improve = feedback.improve.strip()
        helpful = feedback.helpful.strip()
        message = feedback.message.strip()
        category = feedback.category.strip() or "other"
        if not improve and not helpful and not message:
            raise HTTPException(status_code=400, detail="Feedback response is required.")
        feedback_payload = {
            "source": "web",
            "category": category,
            "improve": improve,
            "helpful": helpful,
            "message": message,
            "user_agent": request.headers.get("User-Agent", ""),
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "context": _feedback_request_context(
                request,
                feedback,
                hosting=app.state.hosting_settings,
                stage_store=app.state.stage_store,
            ),
        }
        record_feedback(feedback_payload, output_dir=_feedback_output_dir(app.state.hosting_settings))
        email_sent = send_feedback_email(feedback_payload)
        return {"status": "ok", "emailSent": email_sent}

    @app.get("/api/queue")
    async def queue_snapshot() -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        return {"items": stage_store.snapshot(), "maxFiles": MAX_FILES_PER_BATCH}

    @app.post("/api/upload")
    async def upload(
        files: list[UploadFile] = File(...),
        source_hint: str = Form("upload"),
        filename_style: str = Form(DEFAULT_FILENAME_STYLE),
    ) -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        event_broker: EventBroker = app.state.event_broker
        queued_before = stage_store.active_count()
        normalized_source = "outlook" if source_hint == "outlook" else "upload"
        staged = await stage_store.stage_uploads(files, source=normalized_source)
        resolved_filename_style = normalize_filename_style(filename_style)
        await asyncio.to_thread(stage_store.refresh_output_previews, filename_style=resolved_filename_style)
        batch_meta = build_batch_meta(staged)
        for item in staged:
            meta = batch_meta[item.path.resolve()]
            emit_task_event(publish_task_event, task_id=item.task_id, stage="drop_received", file_name=item.name, meta=meta)
            emit_task_event(publish_task_event, task_id=item.task_id, stage="files_accepted", file_name=item.name, meta=meta)
        if staged:
            event_broker.publish_status(kind="status", message=f"Queued {len(staged)} .msg file(s).")
        rejected_count = max(0, len(files) - len(staged))
        if rejected_count:
            remaining = max(0, MAX_FILES_PER_BATCH - queued_before)
            event_broker.publish_status(
                kind="status",
                message=(
                    f"Accepted {len(staged)} file(s); {rejected_count} were skipped. "
                    f"Queue capacity is {MAX_FILES_PER_BATCH} and {remaining} slot(s) were open."
                ),
            )
        return {
            "items": stage_store.snapshot(),
            "accepted": [item.to_dict() for item in staged],
            "rejectedCount": rejected_count,
        }

    @app.post("/api/filename-style-preview")
    async def filename_style_preview(request: FilenameStyleRequest) -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        resolved_filename_style = normalize_filename_style(request.filename_style)
        await asyncio.to_thread(stage_store.refresh_output_previews, filename_style=resolved_filename_style)
        return {"items": stage_store.snapshot(), "filenameStyle": resolved_filename_style}

    @app.post("/api/import-outlook")
    async def import_outlook() -> dict[str, object]:
        hosting: HostingSettings = app.state.hosting_settings
        if hosting.disable_outlook_import:
            raise HTTPException(status_code=403, detail="Outlook import is disabled for this hosted deployment.")
        stage_store: StageStore = app.state.stage_store
        event_broker: EventBroker = app.state.event_broker
        staged_paths = await asyncio.to_thread(extract_selected_outlook_messages, stage_store.remaining_slots())
        staged = stage_store.stage_paths(staged_paths, source="outlook")
        await asyncio.to_thread(stage_store.refresh_output_previews)
        batch_meta = build_batch_meta(staged)
        for item in staged:
            meta = batch_meta[item.path.resolve()]
            emit_task_event(publish_task_event, task_id=item.task_id, stage="outlook_extract_started", file_name=item.name, meta=meta)
            emit_task_event(publish_task_event, task_id=item.task_id, stage="files_accepted", file_name=item.name, meta=meta)
        if staged:
            event_broker.publish_status(kind="status", message=f"Imported {len(staged)} message(s) from Outlook.")
        return {"items": stage_store.snapshot(), "accepted": [item.to_dict() for item in staged]}

    @app.post("/api/remove")
    async def remove_items(request: RemoveRequest) -> dict[str, object]:
        app.state.stage_store.remove(request.ids)
        return {"items": app.state.stage_store.snapshot()}

    @app.post("/api/clear")
    async def clear_items() -> dict[str, object]:
        app.state.stage_store.clear()
        return {"items": []}

    @app.post("/api/choose-output-folder")
    async def choose_output_folder() -> JSONResponse:
        hosting: HostingSettings = app.state.hosting_settings
        if hosting.default_output_dir is not None:
            output_dir = hosting.default_output_dir
            output_dir.mkdir(parents=True, exist_ok=True)
            return JSONResponse({"outputDir": str(output_dir), "outputDirLabel": output_dir.name or str(output_dir)}, status_code=200)
        if hosting.disable_output_picker:
            return JSONResponse(
                {
                    "outputDir": None,
                    "disabled": True,
                    "reason": "The hosted deployment uses a server-managed output location.",
                },
                status_code=200,
            )
        selected = await asyncio.to_thread(choose_output_directory)
        if not selected:
            return JSONResponse({"outputDir": None}, status_code=200)
        output_dir = Path(selected).expanduser()
        return JSONResponse({"outputDir": str(output_dir), "outputDirLabel": output_dir.name or str(output_dir)})

    @app.post("/api/open-output-folder")
    async def open_output_folder(request: OpenOutputFolderRequest) -> dict[str, object]:
        output_dir = (request.output_dir or "").strip()
        if not output_dir:
            raise HTTPException(status_code=400, detail="Choose an output folder before opening it.")
        opened = await asyncio.to_thread(open_output_directory, output_dir)
        if not opened:
            raise HTTPException(status_code=404, detail="The selected output folder is unavailable.")
        return {"ok": True, "outputDir": output_dir}

    @app.post("/api/reveal-output-file")
    async def reveal_output_file_endpoint(request: RevealOutputFileRequest) -> dict[str, object]:
        output_path = (request.output_path or "").strip()
        if not output_path:
            raise HTTPException(status_code=400, detail="Choose a saved PDF before opening it.")
        opened = await asyncio.to_thread(reveal_output_file, output_path)
        if not opened:
            raise HTTPException(status_code=404, detail="The selected PDF is unavailable.")
        return {"ok": True, "outputPath": output_path}

    @app.post("/api/preview-mailroom")
    async def preview_mailroom(request: PreviewRequest) -> dict[str, object]:
        event_broker: EventBroker = app.state.event_broker
        pipeline = request.pipeline if request.pipeline in {"outlook_edge", "edge_html", "reportlab"} else "outlook_edge"
        asyncio.create_task(publish_preview_sequence(event_broker, pipeline=pipeline))
        return {"ok": True, "pipeline": pipeline}

    @app.post("/api/convert")
    async def convert(request: ConvertRequest) -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        event_broker: EventBroker = app.state.event_broker
        hosting: HostingSettings = app.state.hosting_settings
        items = stage_store.resolve_convertible_items(request.ids)
        output_dir = _resolve_output_dir(request.output_dir, hosting)

        batch_meta = build_batch_meta(items)
        for item in items:
            emit_task_event(
                publish_task_event,
                task_id=item.task_id,
                stage="output_folder_selected",
                file_name=item.name,
                meta={
                    **batch_meta[item.path.resolve()],
                    "outputDir": str(output_dir),
                    "outputDirLabel": output_dir.name or str(output_dir),
                },
            )

        task_ids_by_source_path = {item.path.resolve(): item.task_id for item in items}
        result = await asyncio.to_thread(
            convert_msg_files,
            [item.path for item in items],
            output_dir,
            event_sink=publish_task_event,
            task_ids_by_source_path=task_ids_by_source_path,
            batch_meta_by_source_path=batch_meta,
            filename_style=normalize_filename_style(request.filename_style),
        )
        if result.converted_files:
            event_broker.publish_status(kind="status", message=f"Converted {len(result.converted_files)} of {result.requested_count} file(s).")
        if result.errors:
            event_broker.publish_status(kind="status", message=result.errors[0])
        return summarize_result(result)

    @app.get("/api/events")
    async def event_stream() -> StreamingResponse:
        event_broker: EventBroker = app.state.event_broker
        subscriber = event_broker.subscribe()

        async def stream() -> Any:
            try:
                yield "retry: 1500\n\n"
                while True:
                    try:
                        payload = await asyncio.to_thread(subscriber.get, True, 15)
                        yield f"data: {payload}\n\n"
                    except queue.Empty:
                        yield ": keep-alive\n\n"
            finally:
                event_broker.unsubscribe(subscriber)

        return StreamingResponse(stream(), media_type="text/event-stream")

    return app


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the MSG to PDF browser app.")
    parser.add_argument("--host", default=os.getenv("MSG_TO_PDF_HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.getenv("MSG_TO_PDF_PORT", "8765")))
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    server_settings = load_server_settings(host=args.host, port=args.port)
    validate_startup_contract(server_settings)

    scheme = "https" if server_settings.tls_enabled else "http"
    url = f"{scheme}://{server_settings.host}:{server_settings.port}"
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    uvicorn.run(
        create_app(),
        host=server_settings.host,
        port=server_settings.port,
        log_level="info",
        ssl_certfile=str(server_settings.tls_certfile) if server_settings.tls_certfile else None,
        ssl_keyfile=str(server_settings.tls_keyfile) if server_settings.tls_keyfile else None,
        server_header=False,
    )


if __name__ == "__main__":
    main()
