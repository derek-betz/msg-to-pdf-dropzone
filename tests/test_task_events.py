from __future__ import annotations

import json
from pathlib import Path

from msg_to_pdf_dropzone.task_events import (
    JsonlTaskEventSink,
    TaskEvent,
    build_batch_meta_for_paths,
    build_task_event_sink_from_env,
    default_task_id_for_path,
    emit_task_event,
    merge_event_meta,
)


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


def test_merge_event_meta_returns_none_when_both_empty() -> None:
    assert merge_event_meta(None, None) is None
    assert merge_event_meta({}, {}) is None


def test_merge_event_meta_returns_base_when_extra_is_none() -> None:
    result = merge_event_meta({"a": 1}, None)
    assert result == {"a": 1}


def test_merge_event_meta_combines_both_dicts() -> None:
    result = merge_event_meta({"a": 1}, {"b": 2})
    assert result == {"a": 1, "b": 2}


def test_merge_event_meta_extra_overrides_base_keys() -> None:
    result = merge_event_meta({"a": 1, "b": 2}, {"b": 99})
    assert result == {"a": 1, "b": 99}


def test_build_batch_meta_for_paths_assigns_consistent_metadata(tmp_path: Path) -> None:
    paths = [tmp_path / "a.msg", tmp_path / "b.msg", tmp_path / "c.msg"]
    meta = build_batch_meta_for_paths(paths)

    assert len(meta) == 3
    for path in paths:
        entry = meta[path.resolve()]
        assert entry["batchSize"] == 3
        assert str(entry["batchId"]).startswith("msg-batch-")
    assert meta[paths[0].resolve()]["batchIndex"] == 1
    assert meta[paths[1].resolve()]["batchIndex"] == 2
    assert meta[paths[2].resolve()]["batchIndex"] == 3


def test_build_batch_meta_for_paths_generates_unique_batch_ids(tmp_path: Path) -> None:
    paths = [tmp_path / "x.msg"]
    meta_a = build_batch_meta_for_paths(paths)
    meta_b = build_batch_meta_for_paths(paths)
    assert meta_a[paths[0].resolve()]["batchId"] != meta_b[paths[0].resolve()]["batchId"]


def test_build_task_event_sink_from_env_returns_none_without_env(monkeypatch) -> None:
    monkeypatch.delenv("MSG_TO_PDF_TASK_EVENT_LOG", raising=False)
    assert build_task_event_sink_from_env() is None


def test_build_task_event_sink_from_env_returns_jsonl_sink_when_env_set(monkeypatch, tmp_path: Path) -> None:
    log_path = tmp_path / "events.jsonl"
    monkeypatch.setenv("MSG_TO_PDF_TASK_EVENT_LOG", str(log_path))
    sink = build_task_event_sink_from_env()
    assert sink is not None
    assert isinstance(sink, JsonlTaskEventSink)


def test_task_event_to_dict_omits_none_optional_fields() -> None:
    event = TaskEvent(
        task_id="t1",
        task_type="msg-to-pdf",
        stage="complete",
        timestamp="2026-01-01T00:00:00+00:00",
    )
    payload = event.to_dict()
    assert payload["taskId"] == "t1"
    assert payload["taskType"] == "msg-to-pdf"
    assert payload["stage"] == "complete"
    assert "fileName" not in payload
    assert "pipeline" not in payload
    assert "success" not in payload
    assert "error" not in payload
    assert "meta" not in payload


def test_task_event_to_dict_includes_all_set_fields() -> None:
    event = TaskEvent(
        task_id="t2",
        task_type="msg-to-pdf",
        stage="pdf_written",
        timestamp="2026-01-01T00:00:00+00:00",
        file_name="sample.msg",
        pipeline="reportlab",
        success=True,
        error=None,
        meta={"batchId": "b1", "batchSize": 2},
    )
    payload = event.to_dict()
    assert payload["fileName"] == "sample.msg"
    assert payload["pipeline"] == "reportlab"
    assert payload["success"] is True
    assert "error" not in payload
    assert payload["meta"] == {"batchId": "b1", "batchSize": 2}


def test_emit_task_event_returns_event_without_sink() -> None:
    event = emit_task_event(
        None,
        task_id="task-no-sink",
        stage="parse_started",
        file_name="file.msg",
    )
    assert event.task_id == "task-no-sink"
    assert event.stage == "parse_started"
    assert event.file_name == "file.msg"
