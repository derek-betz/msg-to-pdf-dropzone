from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import threading
from typing import Literal, Protocol, TypeAlias

TaskType: TypeAlias = Literal["msg-to-pdf"]
TaskStage: TypeAlias = Literal[
    "drop_received",
    "outlook_extract_started",
    "files_accepted",
    "output_folder_selected",
    "parse_started",
    "filename_built",
    "pdf_pipeline_started",
    "pipeline_selected",
    "pdf_written",
    "deliver_started",
    "complete",
    "failed",
]
RenderPipeline: TypeAlias = Literal["outlook_edge", "edge_html", "reportlab"]
TaskMetaValue: TypeAlias = str | int | float | bool | None


@dataclass(slots=True)
class TaskEvent:
    task_id: str
    task_type: TaskType
    stage: TaskStage
    timestamp: str
    file_name: str | None = None
    pipeline: RenderPipeline | None = None
    success: bool | None = None
    error: str | None = None
    meta: dict[str, TaskMetaValue] | None = None

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {
            "taskId": self.task_id,
            "taskType": self.task_type,
            "stage": self.stage,
            "timestamp": self.timestamp,
        }
        if self.file_name is not None:
            payload["fileName"] = self.file_name
        if self.pipeline is not None:
            payload["pipeline"] = self.pipeline
        if self.success is not None:
            payload["success"] = self.success
        if self.error is not None:
            payload["error"] = self.error
        if self.meta:
            payload["meta"] = dict(self.meta)
        return payload


class TaskEventSink(Protocol):
    def __call__(self, event: TaskEvent) -> None:
        ...


class JsonlTaskEventSink:
    def __init__(self, output_path: Path) -> None:
        self.output_path = output_path
        self._lock = threading.Lock()

    def __call__(self, event: TaskEvent) -> None:
        line = json.dumps(event.to_dict(), ensure_ascii=True, sort_keys=True)
        with self._lock:
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            with self.output_path.open("a", encoding="utf-8") as handle:
                handle.write(line)
                handle.write("\n")


def current_event_timestamp() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def default_task_id_for_path(path: Path) -> str:
    normalized = str(path.expanduser().resolve()).lower()
    digest = hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:12]
    return f"msg-to-pdf-{digest}"


def emit_task_event(
    sink: TaskEventSink | None,
    *,
    task_id: str,
    stage: TaskStage,
    task_type: TaskType = "msg-to-pdf",
    timestamp: str | None = None,
    file_name: str | None = None,
    pipeline: RenderPipeline | None = None,
    success: bool | None = None,
    error: str | None = None,
    meta: Mapping[str, TaskMetaValue] | None = None,
) -> TaskEvent:
    event = TaskEvent(
        task_id=task_id,
        task_type=task_type,
        stage=stage,
        timestamp=timestamp or current_event_timestamp(),
        file_name=file_name,
        pipeline=pipeline,
        success=success,
        error=error,
        meta=dict(meta) if meta else None,
    )
    if sink is not None:
        sink(event)
    return event


def build_task_event_sink_from_env() -> TaskEventSink | None:
    raw_path = os.environ.get("MSG_TO_PDF_TASK_EVENT_LOG", "").strip()
    if not raw_path:
        return None
    return JsonlTaskEventSink(Path(raw_path).expanduser())
