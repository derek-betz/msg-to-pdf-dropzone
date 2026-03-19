from __future__ import annotations

from pathlib import Path

from msg_to_pdf_dropzone.app_state import AppState, get_app_state_dir, load_app_state, save_app_state


def test_save_and_load_app_state_round_trip(tmp_path: Path) -> None:
    state_path = tmp_path / "state" / "app-state.json"

    save_app_state(AppState(theater_open=True), state_path)
    loaded = load_app_state(state_path)

    assert loaded.theater_open is True


def test_get_app_state_dir_prefers_xdg_config_home(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))

    assert get_app_state_dir() == tmp_path / "msg-to-pdf-dropzone"
