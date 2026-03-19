from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime, timedelta
from functools import partial
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import tempfile
import threading
import time
from types import ModuleType
from typing import Any
from urllib.parse import parse_qs, urlparse
import webbrowser

from .app_state import AppState, load_app_state, save_app_state
from .task_events import JsonlTaskEventSink, TaskEventSink

THEATER_ENABLE_ENV_VAR = "MSG_TO_PDF_ENABLE_THEATER"
THEATER_EVENT_LOG_ENV_VAR = "MSG_TO_PDF_TASK_EVENT_LOG"
THEATER_TITLE = "MSG to PDF Theater"
MAX_DEMO_FILE_COUNT = 10


def get_theater_assets_dir() -> Path:
    return Path(__file__).resolve().parent / "theater_assets"


def theater_assets_available(assets_dir: Path | None = None) -> bool:
    resolved_assets_dir = assets_dir or get_theater_assets_dir()
    return (resolved_assets_dir / "index.html").exists()


def _env_flag_is_enabled(name: str) -> bool:
    raw_value = os.environ.get(name, "").strip().lower()
    return raw_value in {"1", "true", "yes", "on"}


def _default_theater_event_log_path() -> Path:
    temp_dir = Path(tempfile.gettempdir())
    return temp_dir / "msg-to-pdf-dropzone-task-events.jsonl"


def resolve_theater_event_log_path() -> Path | None:
    raw_path = os.environ.get(THEATER_EVENT_LOG_ENV_VAR, "").strip()
    if raw_path:
        return Path(raw_path).expanduser()
    if _env_flag_is_enabled(THEATER_ENABLE_ENV_VAR):
        return _default_theater_event_log_path()
    return None


def read_event_payloads_from_jsonl(log_path: Path) -> list[dict[str, object]]:
    try:
        lines = log_path.read_text(encoding="utf-8").splitlines()
    except FileNotFoundError:
        return []

    payloads: list[dict[str, object]] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        try:
            payload = json.loads(stripped)
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict):
            payloads.append(payload)
    return payloads


class _TheaterRequestHandler(SimpleHTTPRequestHandler):
    def __init__(
        self,
        *args: object,
        directory: str,
        theater_server: "TheaterServer",
        **kwargs: object,
    ) -> None:
        self._theater_server = theater_server
        super().__init__(*args, directory=directory, **kwargs)

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/health":
            self._send_json({"ok": True})
            return
        if parsed.path == "/api/events":
            after = _parse_after_offset(parse_qs(parsed.query))
            events = read_event_payloads_from_jsonl(self._theater_server.event_log_path)
            self._send_json(
                {
                    "events": events[after:],
                    "nextOffset": len(events),
                }
            )
            return
        super().do_GET()

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path != "/api/demo":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        payload = self._read_json_body()
        file_count = _coerce_demo_count(payload.get("fileCount"), default=5)
        failure_count = _coerce_demo_count(
            payload.get("failureCount"),
            default=0,
            max_value=file_count,
        )
        started = self._theater_server.start_demo(
            file_count=file_count,
            failure_count=failure_count,
        )
        self._send_json({"ok": True, **started})

    def log_message(self, _format: str, *_args: object) -> None:
        return

    def _send_json(self, payload: dict[str, object]) -> None:
        body = json.dumps(payload, ensure_ascii=True).encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json_body(self) -> dict[str, object]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except Exception:
            length = 0
        raw = self.rfile.read(max(0, length))
        if not raw:
            return {}
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception:
            return {}
        return payload if isinstance(payload, dict) else {}


def _parse_after_offset(values: dict[str, list[str]]) -> int:
    raw = values.get("after", ["0"])[0]
    try:
        return max(0, int(raw))
    except Exception:
        return 0


def _coerce_demo_count(
    value: object,
    *,
    default: int,
    min_value: int = 0,
    max_value: int = MAX_DEMO_FILE_COUNT,
) -> int:
    try:
        coerced = int(value)
    except Exception:
        coerced = default
    return max(min_value, min(max_value, coerced))


class TheaterServer:
    def __init__(
        self,
        *,
        assets_dir: Path,
        event_log_path: Path,
        demo_initial_delay_s: float = 0.9,
        demo_step_scale: float = 1.0,
    ) -> None:
        self.assets_dir = assets_dir
        self.event_log_path = event_log_path
        self._httpd: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None
        self.base_url: str | None = None
        self._demo_lock = threading.Lock()
        self._demo_generation = 0
        self._demo_initial_delay_s = demo_initial_delay_s
        self._demo_step_scale = demo_step_scale

    def start(self) -> str:
        if self.base_url is not None:
            return self.base_url
        handler = partial(
            _TheaterRequestHandler,
            directory=str(self.assets_dir),
            theater_server=self,
        )
        httpd = ThreadingHTTPServer(("127.0.0.1", 0), handler)
        thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        thread.start()
        self._httpd = httpd
        self._thread = thread
        host, port = httpd.server_address
        self.base_url = f"http://{host}:{port}/"
        return self.base_url

    def stop(self) -> None:
        if self._httpd is not None:
            self._httpd.shutdown()
            self._httpd.server_close()
        if self._thread is not None:
            self._thread.join(timeout=2.0)
        with self._demo_lock:
            self._demo_generation += 1
        self._httpd = None
        self._thread = None
        self.base_url = None

    def start_demo(self, *, file_count: int, failure_count: int) -> dict[str, int]:
        normalized_file_count = _coerce_demo_count(file_count, default=5, min_value=1)
        normalized_failure_count = _coerce_demo_count(
            failure_count,
            default=0,
            max_value=normalized_file_count,
        )
        with self._demo_lock:
            self._demo_generation += 1
            generation = self._demo_generation
            self.event_log_path.parent.mkdir(parents=True, exist_ok=True)
            self.event_log_path.write_text("", encoding="utf-8")

        thread = threading.Thread(
            target=self._run_demo,
            args=(generation, normalized_file_count, normalized_failure_count),
            daemon=True,
        )
        thread.start()
        return {
            "fileCount": normalized_file_count,
            "failureCount": normalized_failure_count,
        }

    def _run_demo(
        self,
        generation: int,
        file_count: int,
        failure_count: int,
    ) -> None:
        if not self._sleep_step(generation, self._demo_initial_delay_s):
            return

        batch_id = f"demo-batch-{int(time.time() * 1000)}"
        base_timestamp = datetime.now(UTC)
        offset_seconds = 0
        successes_before_failure = file_count - failure_count
        names = [
            "alpha",
            "bravo",
            "charlie",
            "delta",
            "echo",
            "foxtrot",
            "golf",
            "hotel",
            "india",
            "juliet",
        ]
        pipelines = ["outlook_edge", "edge_html", "reportlab"]

        sequence: list[tuple[float, dict[str, object]]] = []
        for index in range(1, file_count + 1):
            stem = names[index - 1] if index <= len(names) else f"file-{index:02d}"
            file_name = f"{stem}.msg"
            output_name = f"2026-03-19_{stem.title().replace('-', '_')}.pdf"
            pipeline = pipelines[(index - 1) % len(pipelines)]
            meta = {
                "batchId": batch_id,
                "batchSize": file_count,
                "batchIndex": index,
                "outputDirLabel": "Demo Destination",
                "outputName": output_name,
            }
            task_id = f"{batch_id}-task-{index}"
            is_failure = index > successes_before_failure

            sequence.extend(
                [
                    (
                        0.0 if not sequence else 0.65,
                        self._build_demo_event(
                            task_id,
                            "drop_received",
                            base_timestamp,
                            offset_seconds,
                            fileName=file_name,
                        ),
                    ),
                    (
                        0.85,
                        self._build_demo_event(
                            task_id,
                            "files_accepted",
                            base_timestamp,
                            offset_seconds + 1,
                            fileName=file_name,
                            meta=meta,
                        ),
                    ),
                    (
                        1.0 if index == 1 else 0.82,
                        self._build_demo_event(
                            task_id,
                            "output_folder_selected",
                            base_timestamp,
                            offset_seconds + 2,
                            fileName=file_name,
                            meta=meta,
                        ),
                    ),
                    (
                        0.8,
                        self._build_demo_event(
                            task_id,
                            "parse_started",
                            base_timestamp,
                            offset_seconds + 3,
                            fileName=file_name,
                            meta=meta,
                        ),
                    ),
                    (
                        0.82,
                        self._build_demo_event(
                            task_id,
                            "filename_built",
                            base_timestamp,
                            offset_seconds + 4,
                            fileName=file_name,
                            meta=meta,
                        ),
                    ),
                    (
                        0.9,
                        self._build_demo_event(
                            task_id,
                            "pdf_pipeline_started",
                            base_timestamp,
                            offset_seconds + 5,
                            fileName=file_name,
                            meta=meta,
                        ),
                    ),
                    (
                        0.72,
                        self._build_demo_event(
                            task_id,
                            "pipeline_selected",
                            base_timestamp,
                            offset_seconds + 6,
                            fileName=file_name,
                            pipeline=pipeline,
                            meta=meta,
                        ),
                    ),
                ]
            )

            if is_failure:
                sequence.append(
                    (
                        1.0,
                        self._build_demo_event(
                            task_id,
                            "failed",
                            base_timestamp,
                            offset_seconds + 7,
                            fileName=file_name,
                            pipeline=pipeline,
                            error="Demo conversion failed",
                            success=False,
                            meta=meta,
                        ),
                    )
                )
            else:
                sequence.extend(
                    [
                        (
                            1.05,
                            self._build_demo_event(
                                task_id,
                                "pdf_written",
                                base_timestamp,
                                offset_seconds + 7,
                                fileName=file_name,
                                pipeline=pipeline,
                                meta=meta,
                            ),
                        ),
                        (
                            0.84,
                            self._build_demo_event(
                                task_id,
                                "deliver_started",
                                base_timestamp,
                                offset_seconds + 8,
                                fileName=file_name,
                                pipeline=pipeline,
                                meta=meta,
                            ),
                        ),
                        (
                            0.94,
                            self._build_demo_event(
                                task_id,
                                "complete",
                                base_timestamp,
                                offset_seconds + 9,
                                fileName=file_name,
                                pipeline=pipeline,
                                success=True,
                                meta=meta,
                            ),
                        ),
                    ]
                )
            offset_seconds += 10

        for delay_s, payload in sequence:
            if not self._sleep_step(generation, delay_s * self._demo_step_scale):
                return
            if not self._emit_demo_event(generation, payload):
                return

    def _sleep_step(self, generation: int, delay_s: float) -> bool:
        deadline = time.monotonic() + max(0.0, delay_s)
        while time.monotonic() < deadline:
            with self._demo_lock:
                if generation != self._demo_generation:
                    return False
            time.sleep(0.02)
        with self._demo_lock:
            return generation == self._demo_generation

    def _emit_demo_event(self, generation: int, payload: dict[str, object]) -> bool:
        with self._demo_lock:
            if generation != self._demo_generation:
                return False
            with self.event_log_path.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(payload, ensure_ascii=True))
                handle.write("\n")
        return True

    def _build_demo_event(
        self,
        task_id: str,
        stage: str,
        base_timestamp: datetime,
        offset_seconds: int,
        **extra: object,
    ) -> dict[str, object]:
        payload = {
            "taskId": task_id,
            "taskType": "msg-to-pdf",
            "stage": stage,
            "timestamp": (base_timestamp + timedelta(seconds=offset_seconds)).isoformat(),
        }
        payload.update(extra)
        return payload


class TheaterController:
    def __init__(
        self,
        *,
        state_path: Path | None = None,
        event_log_path: Path | None = None,
        assets_dir: Path | None = None,
        browser_opener: Callable[[str], bool] | None = None,
        webview_module: ModuleType | None = None,
    ) -> None:
        self._state_path = state_path
        self._state = load_app_state(state_path)
        self._assets_dir = assets_dir or get_theater_assets_dir()
        self._event_log_path = event_log_path or resolve_theater_event_log_path()
        self._browser_opener = browser_opener or webbrowser.open
        self._webview_module = webview_module
        self._server: TheaterServer | None = None
        self._window: Any | None = None
        self._window_thread: threading.Thread | None = None
        self._is_open = False
        self._lock = threading.Lock()
        self._browser_fallback = False
        self._event_sink: TaskEventSink | None = None

        if self._event_log_path is not None:
            self._event_log_path.parent.mkdir(parents=True, exist_ok=True)
            self._event_log_path.write_text("", encoding="utf-8")
            self._event_sink = JsonlTaskEventSink(self._event_log_path)

    @property
    def event_sink(self) -> TaskEventSink | None:
        return self._event_sink

    @property
    def is_open(self) -> bool:
        return self._is_open

    def should_open_on_launch(self) -> bool:
        return _env_flag_is_enabled(THEATER_ENABLE_ENV_VAR) or self._state.theater_open

    def set_persisted_open(self, value: bool) -> None:
        self._state.theater_open = value
        save_app_state(self._state, self._state_path)

    def open(self) -> bool:
        with self._lock:
            if self._is_open:
                return True
            if not theater_assets_available(self._assets_dir):
                return False
            if self._event_log_path is None:
                self._event_log_path = _default_theater_event_log_path()
                self._event_log_path.parent.mkdir(parents=True, exist_ok=True)
                self._event_log_path.write_text("", encoding="utf-8")
                self._event_sink = JsonlTaskEventSink(self._event_log_path)
            if self._server is None:
                self._server = TheaterServer(
                    assets_dir=self._assets_dir,
                    event_log_path=self._event_log_path,
                )
            base_url = self._server.start()

            if os.name == "nt" and self._try_open_pywebview(base_url):
                self._is_open = True
                self._browser_fallback = False
                return True

            self._browser_opener(base_url)
            self._browser_fallback = True
            self._is_open = True
            return True

    def close(self) -> None:
        with self._lock:
            if not self._is_open:
                return
            if self._window is not None:
                try:
                    destroy = getattr(self._window, "destroy", None)
                    hide = getattr(self._window, "hide", None)
                    if callable(hide):
                        hide()
                    elif callable(destroy):
                        destroy()
                except Exception:
                    pass
            if self._browser_fallback and self._server is not None:
                self._server.stop()
                self._server = None
            self._is_open = False
            self._browser_fallback = False

    def shutdown(self) -> None:
        self.close()
        if self._server is not None:
            self._server.stop()
            self._server = None

    def _try_open_pywebview(self, base_url: str) -> bool:
        webview_module = self._resolve_webview_module()
        if webview_module is None:
            return False

        if self._window is not None:
            show = getattr(self._window, "show", None)
            if callable(show):
                try:
                    show()
                    return True
                except Exception:
                    pass

        if self._window_thread is not None and self._window_thread.is_alive():
            return False

        window_ready = threading.Event()

        def _target() -> None:
            try:
                window = webview_module.create_window(
                    THEATER_TITLE,
                    base_url,
                    width=1280,
                    height=860,
                    min_size=(980, 700),
                )
                self._window = window
                window_ready.set()
                webview_module.start()
            except Exception:
                self._window = None
            finally:
                self._is_open = False

        self._window_thread = threading.Thread(target=_target, daemon=True)
        self._window_thread.start()
        window_ready.wait(timeout=3.0)
        return self._window is not None

    def _resolve_webview_module(self) -> ModuleType | None:
        if self._webview_module is not None:
            return self._webview_module
        try:
            import webview  # type: ignore[import-not-found]
        except Exception:
            return None
        self._webview_module = webview
        return self._webview_module
