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
- Microsoft Edge (installed by default on most Windows systems)

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

The app includes an `Open Theater` / `Close Theater` control in the main window. The theater is a companion workshop scene driven by the same normalized task events the converter emits internally.

## Run tests

```powershell
python -m pip install pytest
pytest
```

## Profile sample corpus

Run a local corpus profile against `emails-for-testing` and generate JSON + Markdown summaries:

```powershell
python -m msg_to_pdf_dropzone.corpus_profiler --runs 3 --emails-dir .\emails-for-testing
```

Reports are written under `.local-corpus-profiles\profile-<timestamp>\`.

## Render strategy (local tuning)

- Default (`fidelity`): Outlook MHTML + Edge first, then HTML + Edge, then ReportLab.
- Optional (`fast`): skips Outlook render stage for faster comparisons.

Set strategy for one shell session:

```powershell
$env:MSG_TO_PDF_RENDER_STRATEGY='fast'
```

## Task event log (sprite theater integration)

The app can emit normalized task events to a JSONL file for external visualization tools, including the bundled theater companion.

Set the log path for one shell session:

```powershell
$env:MSG_TO_PDF_TASK_EVENT_LOG='C:\temp\msg-to-pdf-task-events.jsonl'
```

Then run the app normally:

```powershell
python -m msg_to_pdf_dropzone
```

Or enable the bundled theater directly for that shell session:

```powershell
$env:MSG_TO_PDF_ENABLE_THEATER='1'
python -m msg_to_pdf_dropzone
```

When theater mode is enabled, the app creates a local JSONL event stream automatically and opens the companion viewer on launch. On Windows it prefers an embedded `pywebview` window; if that is unavailable it falls back to the default browser.

Each emitted line is one JSON object describing a stage such as:

- `drop_received`
- `outlook_extract_started`
- `files_accepted`
- `output_folder_selected`
- `parse_started`
- `filename_built`
- `pdf_pipeline_started`
- `pipeline_selected`
- `pdf_written`
- `deliver_started`
- `complete`
- `failed`

This is intended for integrations like sprite-based task theater, status visualizers, or local debugging tools.

## Theater runtime development

The bundled theater assets are built from the PixiJS runtime under [`theater_runtime`](./theater_runtime).

Rebuild the packaged theater after editing the runtime:

```powershell
cd .\theater_runtime
npm install
npm test
npm run build
```

The build output is written into `src/msg_to_pdf_dropzone/theater_assets` so the Python package can serve it directly.

## Notes

- Maximum input files per batch is 10.
- Filenames are sanitized for Windows.
- If multiple outputs would have the same name, a numeric suffix is added.
- Outlook drag handling uses COM to export selected items to temporary `.msg` files when direct file paths are not provided.
- PDF generation first tries a high-fidelity `.msg` render via Outlook MHTML export + Edge headless print, then an HTML-to-PDF Edge pass, then falls back to the built-in renderer.
