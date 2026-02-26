from __future__ import annotations

import sys
from pathlib import Path

OL_SAVE_AS_MHTML = 10
OL_SAVE_AS_HTML = 5


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
    try:
        pythoncom.CoInitialize()
        initialized = True

        outlook = win32com.client.DispatchEx("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        item = namespace.OpenSharedItem(str(msg_path))

        for save_type in (OL_SAVE_AS_MHTML, OL_SAVE_AS_HTML):
            try:
                item.SaveAs(str(output_web_path), save_type)
                if output_web_path.exists() and output_web_path.stat().st_size > 0:
                    return 0
            except Exception:
                continue
        return 1
    except Exception:
        return 1
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
