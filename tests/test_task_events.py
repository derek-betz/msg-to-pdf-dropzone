from __future__ import annotations

import json
from pathlib import Path

from msg_to_pdf_dropzone.task_events import JsonlTaskEventSink, default_task_id_for_path, emit_task_event


def test_default_task_id_for_path_is_stable(tmp_path: Path) -> None:
    sample_path = tmp_path / "sample.msg"
    first = default_task_id_for_path(sample_path)
    second = default_task_id_for_path(sample_path)

    assert first == second
    assert first.startswith("msg-to-pdf-")


def test_jsonl_sink_writes_event_payload(tmp_path: Path) -> None:
    output_path = tmp_path / "events" / "task-events.jsonl"
    sink = JsonlTaskEventSink(output_path)

    emit_task_event(
        sink,
        task_id="task-789",
        stage="files_accepted",
        file_name="sample.msg",
        meta={"outputDirLabel": "Sandbox"},
    )

    lines = output_path.read_text(encoding="utf-8").splitlines()
    assert len(lines) == 1
    payload = json.loads(lines[0])
    assert payload["taskId"] == "task-789"
    assert payload["stage"] == "files_accepted"
    assert payload["fileName"] == "sample.msg"
    assert payload["meta"]["outputDirLabel"] == "Sandbox"
