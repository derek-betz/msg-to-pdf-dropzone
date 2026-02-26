from __future__ import annotations

import sys
from pathlib import Path

WD_EXPORT_FORMAT_PDF = 17


def convert_html_to_pdf_with_word(html_path: Path, output_path: Path) -> int:
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return 2

    word = None
    document = None
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(html_path))
        document.ExportAsFixedFormat(str(output_path), WD_EXPORT_FORMAT_PDF)
        if output_path.exists() and output_path.stat().st_size > 0:
            return 0
        return 1
    except Exception:
        return 1
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        return 64
    html_path = Path(argv[1]).resolve()
    output_path = Path(argv[2]).resolve()
    if not html_path.exists():
        return 66
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return convert_html_to_pdf_with_word(html_path, output_path)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
