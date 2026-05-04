from __future__ import annotations

from pathlib import Path

from msg_to_pdf_dropzone.drop_helpers import is_supported_msg_candidate, parse_drop_paths, wait_for_materialized_file


def test_parse_drop_paths_handles_braced_and_quoted_values() -> None:
    raw = '{C:\\Temp\\one.msg} "C:\\Temp\\two.msg"'
    parsed = parse_drop_paths(raw, lambda value: value.split())
    assert parsed == [Path("C:\\Temp\\one.msg"), Path("C:\\Temp\\two.msg")]


def test_is_supported_msg_candidate_accepts_msg_and_extensionless() -> None:
    assert is_supported_msg_candidate(Path("one.msg"))
    assert is_supported_msg_candidate(Path("temp-drop"))
    assert not is_supported_msg_candidate(Path("notes.txt"))


def test_parse_drop_paths_returns_empty_list_for_empty_input() -> None:
    assert parse_drop_paths("", lambda s: s.split()) == []


def test_parse_drop_paths_strips_whitespace_from_tokens() -> None:
    raw = "  /tmp/file.msg  "
    parsed = parse_drop_paths(raw, lambda s: s.split())
    assert parsed == [Path("/tmp/file.msg")]


def test_parse_drop_paths_handles_multiple_paths() -> None:
    raw = "/tmp/a.msg /tmp/b.msg"
    parsed = parse_drop_paths(raw, lambda s: s.split())
    assert parsed == [Path("/tmp/a.msg"), Path("/tmp/b.msg")]


def test_parse_drop_paths_falls_back_to_whitespace_split_on_splitlist_error() -> None:
    def bad_splitlist(value: str):
        raise RuntimeError("split failed")

    raw = "/tmp/file.msg"
    parsed = parse_drop_paths(raw, bad_splitlist)
    assert parsed == [Path("/tmp/file.msg")]


def test_wait_for_materialized_file_returns_existing_path_immediately(tmp_path: Path) -> None:
    path = tmp_path / "test.msg"
    path.write_text("x", encoding="utf-8")
    result = wait_for_materialized_file(path)
    assert result == path


def test_wait_for_materialized_file_returns_path_even_if_never_created(tmp_path: Path) -> None:
    path = tmp_path / "missing.msg"
    result = wait_for_materialized_file(path, timeout_seconds=0.05)
    assert result == path
    assert not path.exists()
