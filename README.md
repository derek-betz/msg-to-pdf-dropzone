# MSG to PDF Dropzone

Windows desktop tool to drag and drop up to 10 Outlook `.msg` files and convert each message into one PDF.

## What it does

- Accepts dragged `.msg` files (or manual file selection).
- Supports dragging selected messages directly from Classic Outlook.
- Converts each email to one PDF.
- Names each PDF as:
  - `YYYY-MM-DD_<email subject>.pdf`
  - `YYYY-MM-DD` is the latest email date found among dropped emails in the same thread.
- Prompts you to select the save folder before conversion.

## Requirements

- Windows
- Python 3.10+
- Outlook `.msg` files
- Microsoft Word installed is recommended for richer PDF rendering.

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

## Run

```powershell
python -m msg_to_pdf_dropzone
```

Or:

```powershell
msg-to-pdf-dropzone
```

## Run tests

```powershell
python -m pip install pytest
pytest
```

## Notes

- Maximum input files per batch is 10.
- Filenames are sanitized for Windows.
- If multiple outputs would have the same name, a numeric suffix is added.
- Outlook drag handling uses COM to export selected items to temporary `.msg` files when direct file paths are not provided.
- PDF generation first tries a high-fidelity HTML-to-PDF pass through Word automation, then falls back to the built-in renderer if Word conversion is unavailable.
