from __future__ import annotations

import time
from pathlib import Path
from types import SimpleNamespace

import msg_to_pdf_dropzone.app as app_module


class _DummyStatusVar:
    def __init__(self) -> None:
        self.value = ""

    def set(self, value: str) -> None:
        self.value = value

    def get(self) -> str:
        return self.value


class _DummyDropZone:
    def configure(self, **_: object) -> None:
        return


class _DummyListbox:
    def __init__(self) -> None:
        self.items: list[str] = []

    def delete(self, _start: int, _end: object) -> None:
        self.items.clear()

    def insert(self, _index: object, value: str) -> None:
        self.items.append(value)

    def configure(self, **_: object) -> None:
        return


class _DummyRoot:
    def __init__(self) -> None:
        self.tk = SimpleNamespace(splitlist=self._splitlist)

    @staticmethod
    def _splitlist(value: str) -> list[str]:
        return value.split()

    def after(self, _ms: int, _callback: object) -> str:
        return "job-id"

    def after_cancel(self, _job_id: str) -> None:
        return

    def configure(self, **_: object) -> None:
        return


def _build_app_for_drop_tests() -> app_module.MsgToPdfApp:
    app = object.__new__(app_module.MsgToPdfApp)
    app._busy = False
    app.root = _DummyRoot()
    app.status_var = _DummyStatusVar()
    app.drop_zone = _DummyDropZone()
    app.drop_zone_default_bg = "#f3f5f8"
    app.selected_files = []
    app.temp_outlook_files = set()
    app.file_listbox = _DummyListbox()
    app._drop_dispatch_seconds = 0.0
    app._drop_slow_warning_count = 0
    app._last_operation_max_stall_seconds = 0.0
    return app


def test_on_drop_outlook_tokens_is_non_blocking(monkeypatch) -> None:
    app = _build_app_for_drop_tests()
    timeouts: list[float] = []
    background_calls: list[str] = []

    def fake_wait(path: Path, timeout_seconds: float = 2.0) -> Path:
        timeouts.append(timeout_seconds)
        time.sleep(timeout_seconds)
        return path

    monkeypatch.setattr(app_module, "wait_for_materialized_file", fake_wait)
    app._run_in_background = lambda **_: background_calls.append("called") or True

    raw_data = " ".join(f"{{C:\\\\Temp\\\\token-{index}}}" for index in range(30))
    event = SimpleNamespace(data=raw_data, sourcetypes=("FileGroupDescriptorW",))

    started_at = time.perf_counter()
    result = app_module.MsgToPdfApp._on_drop(app, event)
    elapsed = time.perf_counter() - started_at

    assert result == app_module.COPY
    assert background_calls == ["called"]
    assert all(timeout == 0.0 for timeout in timeouts)
    assert elapsed < app_module.DROP_UI_READY_TARGET_SECONDS
    assert app._drop_dispatch_seconds < app_module.DROP_UI_READY_TARGET_SECONDS
    assert app._drop_slow_warning_count == 0


def test_on_drop_non_outlook_uses_small_timeout(monkeypatch) -> None:
    app = _build_app_for_drop_tests()
    timeouts: list[float] = []
    app._run_in_background = lambda **_: True

    def fake_wait(path: Path, timeout_seconds: float = 2.0) -> Path:
        timeouts.append(timeout_seconds)
        return path

    monkeypatch.setattr(app_module, "wait_for_materialized_file", fake_wait)

    event = SimpleNamespace(
        data=r"{C:\\Temp\\one.msg} {C:\\Temp\\two.msg}",
        sourcetypes=("DND_Files",),
    )
    result = app_module.MsgToPdfApp._on_drop(app, event)

    assert result == app_module.COPY
    assert timeouts
    assert all(timeout == app_module.DEFAULT_DROP_MATERIALIZATION_TIMEOUT_SECONDS for timeout in timeouts)
    assert app._drop_dispatch_seconds < app_module.DROP_UI_READY_TARGET_SECONDS


def test_on_drop_updates_selected_list_quickly_for_real_file(tmp_path: Path) -> None:
    app = _build_app_for_drop_tests()
    app._run_in_background = lambda **_: (_ for _ in ()).throw(
        AssertionError("Background Outlook path should not run for normal file drop.")
    )
    msg_path = tmp_path / "sample.msg"
    msg_path.write_text("dummy", encoding="utf-8")

    event = SimpleNamespace(data=f"{{{msg_path}}}", sourcetypes=("DND_Files",))
    result = app_module.MsgToPdfApp._on_drop(app, event)

    assert result == app_module.COPY
    assert app.selected_files == [msg_path.resolve()]
    assert app.file_listbox.items == [str(msg_path.resolve())]
    assert app._drop_dispatch_seconds < app_module.DROP_UI_READY_TARGET_SECONDS


def test_heartbeat_tick_records_stall_for_active_operation() -> None:
    app = _build_app_for_drop_tests()
    app._active_operation_name = "convert"
    app._active_operation_max_stall_seconds = 0.0
    app._heartbeat_last_tick = time.perf_counter() - 0.25

    app_module.MsgToPdfApp._on_heartbeat_tick(app)

    assert app._active_operation_max_stall_seconds > 0.0
    assert app._heartbeat_job_id == "job-id"
