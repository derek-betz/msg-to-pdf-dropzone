from __future__ import annotations

import hashlib
from pathlib import Path
import sys
import threading
import time
from types import ModuleType

import msg_to_pdf_dropzone.outlook_mhtml_worker as worker


def test_export_msg_to_web_archive_uses_temp_copy(monkeypatch, tmp_path: Path) -> None:
    msg_path = tmp_path / "sample.msg"
    msg_path.write_bytes(b"msg-bytes")
    original_hash = hashlib.sha256(msg_path.read_bytes()).hexdigest()
    output_path = tmp_path / "sample.mht"
    opened_paths: list[Path] = []

    class FakeItem:
        def SaveAs(self, destination: str, _save_type: int) -> None:
            Path(destination).write_text("exported", encoding="utf-8")

        def Close(self, _mode: int) -> None:
            return

    class FakeNamespace:
        def OpenSharedItem(self, path: str) -> FakeItem:
            opened_paths.append(Path(path))
            return FakeItem()

    class FakeOutlook:
        def GetNamespace(self, _name: str) -> FakeNamespace:
            return FakeNamespace()

    pythoncom_module = ModuleType("pythoncom")
    pythoncom_module.CoInitialize = lambda: None
    pythoncom_module.CoUninitialize = lambda: None

    client_module = ModuleType("win32com.client")
    client_module.DispatchEx = lambda _name: FakeOutlook()

    win32com_module = ModuleType("win32com")
    win32com_module.client = client_module

    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom_module)
    monkeypatch.setitem(sys.modules, "win32com", win32com_module)
    monkeypatch.setitem(sys.modules, "win32com.client", client_module)

    result = worker.export_msg_to_web_archive(msg_path, output_path)

    assert result == 0
    assert output_path.read_text(encoding="utf-8") == "exported"
    assert opened_paths
    assert opened_paths[0] != msg_path
    assert opened_paths[0].name == msg_path.name
    assert not opened_paths[0].exists()
    assert hashlib.sha256(msg_path.read_bytes()).hexdigest() == original_hash


def test_wait_for_output_file_handles_delayed_write(tmp_path: Path) -> None:
    output_path = tmp_path / "delayed.mht"

    def delayed_write() -> None:
        time.sleep(0.1)
        output_path.write_text("exported", encoding="utf-8")

    writer = threading.Thread(target=delayed_write, daemon=True)
    writer.start()

    assert worker._wait_for_output_file(output_path, timeout_seconds=1.0) is True
    writer.join(timeout=1.0)


def test_export_msg_to_web_archive_accepts_output_materialized_on_close(monkeypatch, tmp_path: Path) -> None:
    msg_path = tmp_path / "sample.msg"
    msg_path.write_bytes(b"msg-bytes")
    output_path = tmp_path / "sample.mht"

    class FakeItem:
        def __init__(self) -> None:
            self.destination = ""

        def SaveAs(self, destination: str, _save_type: int) -> None:
            self.destination = destination

        def Close(self, _mode: int) -> None:
            Path(self.destination).write_text("exported-on-close", encoding="utf-8")

    class FakeNamespace:
        def __init__(self) -> None:
            self.item = FakeItem()

        def OpenSharedItem(self, _path: str) -> FakeItem:
            return self.item

    class FakeOutlook:
        def __init__(self) -> None:
            self.namespace = FakeNamespace()

        def GetNamespace(self, _name: str) -> FakeNamespace:
            return self.namespace

    pythoncom_module = ModuleType("pythoncom")
    pythoncom_module.CoInitialize = lambda: None
    pythoncom_module.CoUninitialize = lambda: None

    client_module = ModuleType("win32com.client")
    client_module.DispatchEx = lambda _name: FakeOutlook()

    win32com_module = ModuleType("win32com")
    win32com_module.client = client_module

    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom_module)
    monkeypatch.setitem(sys.modules, "win32com", win32com_module)
    monkeypatch.setitem(sys.modules, "win32com.client", client_module)

    result = worker.export_msg_to_web_archive(msg_path, output_path)

    assert result == 0
    assert output_path.read_text(encoding="utf-8") == "exported-on-close"
