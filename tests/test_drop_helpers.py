from __future__ import annotations

from pathlib import Path

from msg_to_pdf_dropzone.drop_helpers import is_supported_msg_candidate, parse_drop_paths


def test_parse_drop_paths_handles_braced_and_quoted_values() -> None:
    raw = '{C:\\Temp\\one.msg} "C:\\Temp\\two.msg"'
    parsed = parse_drop_paths(raw, lambda value: value.split())
    assert parsed == [Path("C:\\Temp\\one.msg"), Path("C:\\Temp\\two.msg")]


def test_is_supported_msg_candidate_accepts_msg_and_extensionless() -> None:
    assert is_supported_msg_candidate(Path("one.msg"))
    assert is_supported_msg_candidate(Path("temp-drop"))
    assert not is_supported_msg_candidate(Path("notes.txt"))
