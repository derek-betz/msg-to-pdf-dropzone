from __future__ import annotations

import json
from pathlib import Path
import time
from urllib.request import Request, urlopen

from msg_to_pdf_dropzone.app_state import AppState, save_app_state
from msg_to_pdf_dropzone.theater_host import (
    THEATER_ENABLE_ENV_VAR,
    TheaterController,
    TheaterServer,
    read_event_payloads_from_jsonl,
)


def test_read_event_payloads_from_jsonl_ignores_invalid_lines(tmp_path: Path) -> None:
    log_path = tmp_path / "events.jsonl"
    log_path.write_text(
        '\n{"taskId":"task-1","taskType":"msg-to-pdf","stage":"files_accepted","timestamp":"2026-03-19T12:00:00Z"}\nnot-json\n',
        encoding="utf-8",
    )

    payloads = read_event_payloads_from_jsonl(log_path)

    assert len(payloads) == 1
    assert payloads[0]["taskId"] == "task-1"


def test_theater_server_serves_health_and_event_offsets(tmp_path: Path) -> None:
    assets_dir = tmp_path / "assets"
    assets_dir.mkdir()
    (assets_dir / "index.html").write_text("<!doctype html><title>Theater</title>", encoding="utf-8")

    log_path = tmp_path / "events.jsonl"
    log_path.write_text(
        "\n".join(
            [
                json.dumps(
                    {
                        "taskId": "task-1",
                        "taskType": "msg-to-pdf",
                        "stage": "drop_received",
                        "timestamp": "2026-03-19T12:00:00Z",
                    }
                ),
                json.dumps(
                    {
                        "taskId": "task-1",
                        "taskType": "msg-to-pdf",
                        "stage": "complete",
                        "timestamp": "2026-03-19T12:00:01Z",
                    }
                ),
            ]
        ),
        encoding="utf-8",
    )

    server = TheaterServer(assets_dir=assets_dir, event_log_path=log_path)
    base_url = server.start()

    try:
        with urlopen(f"{base_url}api/health") as response:
            health_payload = json.loads(response.read().decode("utf-8"))
        with urlopen(f"{base_url}api/events?after=1") as response:
            events_payload = json.loads(response.read().decode("utf-8"))
    finally:
        server.stop()

    assert health_payload == {"ok": True}
    assert events_payload["nextOffset"] == 2
    assert len(events_payload["events"]) == 1
    assert events_payload["events"][0]["stage"] == "complete"


def test_theater_server_demo_endpoint_emits_requested_failures(tmp_path: Path) -> None:
    assets_dir = tmp_path / "assets"
    assets_dir.mkdir()
    (assets_dir / "index.html").write_text(
        "<!doctype html><title>Theater</title>",
        encoding="utf-8",
    )

    log_path = tmp_path / "events.jsonl"
    server = TheaterServer(
        assets_dir=assets_dir,
        event_log_path=log_path,
        demo_initial_delay_s=0.01,
        demo_step_scale=0.01,
    )
    base_url = server.start()

    try:
        request = Request(
            f"{base_url}api/demo",
            data=json.dumps({"fileCount": 4, "failureCount": 2}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urlopen(request) as response:
            demo_payload = json.loads(response.read().decode("utf-8"))

        deadline = time.monotonic() + 2.0
        events_payload: dict[str, object] = {}
        while time.monotonic() < deadline:
            with urlopen(f"{base_url}api/events?after=0") as response:
                events_payload = json.loads(response.read().decode("utf-8"))
            events = events_payload.get("events", [])
            if isinstance(events, list):
                terminal_stages = [
                    event["stage"]
                    for event in events
                    if event["stage"] in {"complete", "failed"}
                ]
                if len(terminal_stages) == 4:
                    break
            time.sleep(0.05)
    finally:
        server.stop()

    events = events_payload["events"]
    assert demo_payload["ok"] is True
    assert demo_payload["fileCount"] == 4
    assert demo_payload["failureCount"] == 2
    assert isinstance(events, list)
    assert events[0]["stage"] == "drop_received"

    terminal_stages = [event["stage"] for event in events if event["stage"] in {"complete", "failed"}]
    assert terminal_stages.count("complete") == 2
    assert terminal_stages.count("failed") == 2
    assert terminal_stages[-2:] == ["failed", "failed"]


def test_theater_controller_should_open_on_launch_from_state_or_env(
    monkeypatch,
    tmp_path: Path,
) -> None:
    state_path = tmp_path / "state.json"
    save_app_state(AppState(theater_open=True), state_path)

    controller = TheaterController(
        state_path=state_path,
        assets_dir=tmp_path / "missing-assets",
    )
    assert controller.should_open_on_launch() is True

    monkeypatch.setenv(THEATER_ENABLE_ENV_VAR, "1")
    controller_from_env = TheaterController(
        state_path=tmp_path / "state-2.json",
        assets_dir=tmp_path / "missing-assets-2",
    )
    assert controller_from_env.should_open_on_launch() is True


def test_theater_controller_open_uses_browser_fallback(tmp_path: Path) -> None:
    assets_dir = tmp_path / "assets"
    assets_dir.mkdir()
    (assets_dir / "index.html").write_text("<!doctype html><title>Theater</title>", encoding="utf-8")

    browser_calls: list[str] = []
    controller = TheaterController(
        state_path=tmp_path / "state.json",
        event_log_path=tmp_path / "events.jsonl",
        assets_dir=assets_dir,
        browser_opener=lambda url: browser_calls.append(url) or True,
    )

    try:
        assert controller.open() is True
        assert browser_calls
        assert controller.event_sink is not None

        with urlopen(f"{browser_calls[0]}api/health") as response:
            health_payload = json.loads(response.read().decode("utf-8"))
    finally:
        controller.shutdown()

    assert health_payload == {"ok": True}


def test_theater_controller_open_returns_false_without_assets(tmp_path: Path) -> None:
    controller = TheaterController(
        state_path=tmp_path / "state.json",
        event_log_path=tmp_path / "events.jsonl",
        assets_dir=tmp_path / "missing-assets",
        browser_opener=lambda _url: True,
    )

    assert controller.open() is False
