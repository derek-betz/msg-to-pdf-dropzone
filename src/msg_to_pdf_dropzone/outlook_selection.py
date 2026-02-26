from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Iterable
from uuid import uuid4

from .thread_logic import sanitize_filename_part

OUTLOOK_MARKERS = {
    "filegroupdescriptor",
    "filegroupdescriptorw",
    "renprivatemessages",
    "filecontents",
}


def is_likely_outlook_drop(raw_data: str, sourcetypes: Iterable[str] | object | None) -> bool:
    if isinstance(sourcetypes, (list, tuple, set)):
        values = [str(value).strip().lower() for value in sourcetypes if value]
        if any(marker in value for value in values for marker in OUTLOOK_MARKERS):
            return True

    if raw_data and "outlook" in raw_data.lower():
        return True

    return False


def extract_selected_outlook_messages(max_files: int) -> list[Path]:
    if os.name != "nt" or max_files <= 0:
        return []

    try:
        import pythoncom
        import win32com.client
    except Exception:
        return []

    file_paths: list[Path] = []
    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True

        outlook_app = win32com.client.GetActiveObject("Outlook.Application")
        explorer = outlook_app.ActiveExplorer()
        if explorer is None:
            return []

        selection = explorer.Selection
        if selection is None:
            return []

        count = int(selection.Count)
        if count <= 0:
            return []

        for index in range(1, count + 1):
            if len(file_paths) >= max_files:
                break
            item = selection.Item(index)
            if item is None:
                continue
            output = save_outlook_item_to_temp_msg(item)
            if output is not None:
                file_paths.append(output)
    except Exception:
        return []
    finally:
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return file_paths


def save_outlook_item_to_temp_msg(outlook_item: object) -> Path | None:
    subject = str(getattr(outlook_item, "Subject", "") or "").strip()
    safe_subject = sanitize_filename_part(subject, max_length=80)
    temp_name = f"{uuid4().hex}-{safe_subject}.msg"
    temp_path = Path(tempfile.gettempdir()) / temp_name

    for save_type in (3, 9, None):
        try:
            if save_type is None:
                getattr(outlook_item, "SaveAs")(str(temp_path))
            else:
                getattr(outlook_item, "SaveAs")(str(temp_path), save_type)
        except Exception:
            continue

        if temp_path.exists() and temp_path.stat().st_size > 0:
            return temp_path

    try:
        if temp_path.exists():
            temp_path.unlink()
    except Exception:
        pass

    return None
