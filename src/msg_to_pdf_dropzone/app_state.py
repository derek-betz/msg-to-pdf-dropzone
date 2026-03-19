from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path


APP_STATE_DIR_NAME = "msg-to-pdf-dropzone"
APP_STATE_FILE_NAME = "app-state.json"


@dataclass(slots=True)
class AppState:
    theater_open: bool = False


def get_app_state_dir() -> Path:
    appdata = os.environ.get("APPDATA", "").strip()
    if appdata:
        return Path(appdata) / APP_STATE_DIR_NAME
    if os.name == "nt":
        return Path.home() / "AppData" / "Roaming" / APP_STATE_DIR_NAME
    xdg_config_home = os.environ.get("XDG_CONFIG_HOME", "").strip()
    if xdg_config_home:
        return Path(xdg_config_home) / APP_STATE_DIR_NAME
    return Path.home() / ".config" / APP_STATE_DIR_NAME


def get_app_state_path() -> Path:
    return get_app_state_dir() / APP_STATE_FILE_NAME


def load_app_state(path: Path | None = None) -> AppState:
    state_path = path or get_app_state_path()
    try:
        payload = json.loads(state_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return AppState()
    except Exception:
        return AppState()
    return AppState(theater_open=bool(payload.get("theaterOpen", False)))


def save_app_state(state: AppState, path: Path | None = None) -> None:
    state_path = path or get_app_state_path()
    state_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"theaterOpen": state.theater_open}
    state_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
