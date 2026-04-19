from __future__ import annotations

import argparse
import asyncio
import atexit
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
import queue
import tempfile
import threading
from typing import Any
from uuid import uuid4
import webbrowser

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn

from .converter import MAX_FILES_PER_BATCH, ConversionResult, convert_msg_files
from .outlook_selection import extract_selected_outlook_messages
from .task_events import TaskEvent, TaskMetaValue, build_batch_meta_for_paths, default_task_id_for_path, emit_task_event

PACKAGE_ROOT = Path(__file__).resolve().parent
WEB_UI_DIR = PACKAGE_ROOT / "web_ui"
ASSET_DIR = PACKAGE_ROOT / "assets"
STAGING_DIR = Path(tempfile.gettempdir()) / "msg-to-pdf-browser-staging"


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
        output_path = None
        if event.meta:
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
            if output_path is not None:
                item.output_path = output_path

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
    output_dir: str


class PreviewRequest(BaseModel):
    pipeline: str = "outlook_edge"


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


def build_batch_meta(items: list[StagedFile]) -> dict[Path, dict[str, TaskMetaValue]]:
    return build_batch_meta_for_paths([item.path for item in items])


def summarize_result(result: ConversionResult) -> dict[str, object]:
    return {
        "requestedCount": result.requested_count,
        "convertedFiles": [str(path) for path in result.converted_files],
        "errors": list(result.errors),
        "skippedFiles": [str(path) for path in result.skipped_files],
        "totalSeconds": result.total_seconds,
        "timingLines": list(result.timing_lines),
    }


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
        return {"ok": True, "queuedCount": stage_store.count(), "maxFiles": MAX_FILES_PER_BATCH}

    @app.get("/api/queue")
    async def queue_snapshot() -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        return {"items": stage_store.snapshot(), "maxFiles": MAX_FILES_PER_BATCH}

    @app.post("/api/upload")
    async def upload(files: list[UploadFile] = File(...), source_hint: str = Form("upload")) -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        event_broker: EventBroker = app.state.event_broker
        queued_before = stage_store.active_count()
        normalized_source = "outlook" if source_hint == "outlook" else "upload"
        staged = await stage_store.stage_uploads(files, source=normalized_source)
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

    @app.post("/api/import-outlook")
    async def import_outlook() -> dict[str, object]:
        stage_store: StageStore = app.state.stage_store
        event_broker: EventBroker = app.state.event_broker
        staged_paths = await asyncio.to_thread(extract_selected_outlook_messages, stage_store.remaining_slots())
        staged = stage_store.stage_paths(staged_paths, source="outlook")
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
        selected = await asyncio.to_thread(choose_output_directory)
        if not selected:
            return JSONResponse({"outputDir": None}, status_code=200)
        output_dir = Path(selected).expanduser()
        return JSONResponse({"outputDir": str(output_dir), "outputDirLabel": output_dir.name or str(output_dir)})

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
        items = stage_store.resolve_convertible_items(request.ids)
        output_dir = Path(request.output_dir).expanduser()
        output_dir.mkdir(parents=True, exist_ok=True)

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
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    url = f"http://{args.host}:{args.port}"
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    uvicorn.run(create_app(), host=args.host, port=args.port, log_level="info")
