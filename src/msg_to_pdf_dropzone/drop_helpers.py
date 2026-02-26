from __future__ import annotations

import time
from pathlib import Path
from typing import Callable, Iterable


def parse_drop_paths(raw_data: str, splitlist: Callable[[str], Iterable[str]]) -> list[Path]:
    if not raw_data:
        return []

    try:
        tokens = list(splitlist(raw_data))
    except Exception:
        tokens = raw_data.split()

    results: list[Path] = []
    for token in tokens:
        cleaned = token.strip()
        if cleaned.startswith("{") and cleaned.endswith("}"):
            cleaned = cleaned[1:-1]
        if len(cleaned) >= 2 and cleaned[0] == '"' and cleaned[-1] == '"':
            cleaned = cleaned[1:-1]
        cleaned = cleaned.strip()
        if cleaned:
            results.append(Path(cleaned))
    return results


def is_supported_msg_candidate(path: Path) -> bool:
    return path.suffix.lower() == ".msg" or path.suffix == ""


def wait_for_materialized_file(path: Path, timeout_seconds: float = 2.0) -> Path:
    if path.exists():
        return path

    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if path.exists():
            return path
        time.sleep(0.1)
    return path
