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


def test_get_app_state_dir_uses_appdata_when_set(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert get_app_state_dir() == tmp_path / "msg-to-pdf-dropzone"


def test_load_app_state_returns_default_when_file_is_missing(tmp_path: Path) -> None:
    state = load_app_state(tmp_path / "nonexistent.json")
    assert state.theater_open is False


def test_load_app_state_returns_default_on_corrupt_json(tmp_path: Path) -> None:
    bad_file = tmp_path / "state.json"
    bad_file.write_text("not valid json!", encoding="utf-8")
    state = load_app_state(bad_file)
    assert state.theater_open is False


def test_save_app_state_creates_parent_directories(tmp_path: Path) -> None:
    nested = tmp_path / "a" / "b" / "c" / "state.json"
    save_app_state(AppState(theater_open=False), nested)
    assert nested.exists()


def test_load_app_state_defaults_theater_open_to_false_when_key_missing(tmp_path: Path) -> None:
    state_path = tmp_path / "state.json"
    state_path.write_text("{}", encoding="utf-8")
    state = load_app_state(state_path)
    assert state.theater_open is False
