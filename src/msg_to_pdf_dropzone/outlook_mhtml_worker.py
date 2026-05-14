from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
import tempfile
import time
from uuid import uuid4

OL_SAVE_AS_MHTML = 10
OL_SAVE_AS_HTML = 5
OUTPUT_WAIT_SECONDS = 2.0
OUTPUT_POLL_SECONDS = 0.05


def _wait_for_output_file(output_web_path: Path, *, timeout_seconds: float = OUTPUT_WAIT_SECONDS) -> bool:
    deadline = time.perf_counter() + timeout_seconds
    while time.perf_counter() < deadline:
        try:
            if output_web_path.exists() and output_web_path.stat().st_size > 0:
                return True
        except OSError:
            pass
        time.sleep(OUTPUT_POLL_SECONDS)
    try:
        return output_web_path.exists() and output_web_path.stat().st_size > 0
    except OSError:
        return False


def _writable_temp_root(candidate: Path) -> Path | None:
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        probe_path = candidate / f".write-test-{os.getpid()}-{uuid4().hex[:8]}.tmp"
        probe_path.write_text("ok", encoding="utf-8")
        probe_path.unlink(missing_ok=True)
    except OSError:
        return None
    return candidate


def _candidate_temp_roots() -> list[Path]:
    roots: list[Path] = []
    temp_root = os.environ.get("MSG_TO_PDF_TEMP_DIR", "").strip()
    if temp_root:
        roots.append(Path(temp_root).expanduser())
    roots.append(Path.cwd() / ".msg-to-pdf-temp")
    roots.append(Path(tempfile.gettempdir()) / "msg-to-pdf-dropzone")
    return roots


def _make_temp_dir(prefix: str) -> Path:
    last_error: OSError | None = None
    for root in _candidate_temp_roots():
        try:
            writable_root = _writable_temp_root(root)
            if writable_root is None:
                continue
            for _ in range(100):
                temp_dir = writable_root / f"{prefix}{uuid4().hex[:10]}"
                try:
                    temp_dir.mkdir(parents=False, exist_ok=False)
                except FileExistsError:
                    continue
                if _writable_temp_root(temp_dir) is not None:
                    return temp_dir
                shutil.rmtree(temp_dir, ignore_errors=True)
        except OSError as exc:
            last_error = exc
            continue
    if last_error is not None:
        raise last_error
    raise OSError("Could not create a writable temporary Outlook export directory.")


def export_msg_to_web_archive(msg_path: Path, output_web_path: Path) -> int:
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return 2

    outlook = None
    namespace = None
    item = None
    initialized = False
    result_code = 1
    try:
        temp_dir = _make_temp_dir("msg-to-pdf-outlook-msg-")
        try:
            staged_msg_path = temp_dir / msg_path.name
            shutil.copy2(msg_path, staged_msg_path)

            pythoncom.CoInitialize()
            initialized = True

            outlook = win32com.client.DispatchEx("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            item = namespace.OpenSharedItem(str(staged_msg_path))

            for save_type in (OL_SAVE_AS_MHTML, OL_SAVE_AS_HTML):
                try:
                    item.SaveAs(str(output_web_path), save_type)
                    if _wait_for_output_file(output_web_path):
                        result_code = 0
                        break
                except Exception:
                    continue
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        result_code = 1
    finally:
        if item is not None:
            try:
                item.Close(0)
            except Exception:
                pass
        namespace = None
        outlook = None
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    if result_code != 0 and _wait_for_output_file(output_web_path, timeout_seconds=1.0):
        return 0
    return result_code


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        return 64
    msg_path = Path(argv[1]).resolve()
    output_path = Path(argv[2]).resolve()
    if not msg_path.exists():
        return 66
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return export_msg_to_web_archive(msg_path, output_path)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
