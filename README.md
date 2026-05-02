# MSG to PDF Dropzone

Windows-local browser app for converting Outlook `.msg` files into clean PDFs with the existing Python conversion engine.

## What it does

- Runs a local FastAPI server with a browser UI at `http://127.0.0.1:8765`.
- Accepts dragged `.msg` files, manual browser uploads, or a Classic Outlook selection import.
- Keeps a local queue, a native output-folder chooser, and a live status log.
- Drives an embedded mailroom companion from the real task event stream.
- Converts each email to one PDF named `YYYY-MM-DD_<email subject>.pdf`, using the latest thread date.

## Requirements

- Windows
- Python 3.10+
- Outlook `.msg` files
- Microsoft Edge (installed by default on most Windows systems)

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -e .
```

## Run browser app

```powershell
.\.venv\Scripts\python.exe -m msg_to_pdf_dropzone --port 8765
```

Or:

```powershell
msg-to-pdf-dropzone
```

Useful flags:

- `--port 8765` to choose the local port
- `--host 127.0.0.1` to bind explicitly to localhost
- `--no-browser` to skip auto-opening the browser tab

## Hosted server mode

For `WEB-SVR03`-style hosting, run the app as a loopback-bound background service and set these environment variables:

- `MSG_TO_PDF_SERVER_MODE=1`
- `MSG_TO_PDF_OUTPUT_DIR=<server managed folder>`
- `MSG_TO_PDF_STAGING_DIR=<server writable staging folder>`

In hosted mode, the app can use a fixed server-managed output directory and disable desktop-only features such as native folder selection and Outlook import.

The browser UI is the primary app surface. It includes:

- drag/drop `.msg` upload and manual file add
- Classic Outlook selection import
- native output-folder chooser
- conversion queue controls
- live server-sent task events
- embedded mailroom companion and preview flow

## Optional legacy desktop entrypoint

The old Tk desktop shell still exists for fallback work, but it is no longer the primary UI:

```powershell
msg-to-pdf-desktop
```

## Run tests

```powershell
.\.venv\Scripts\python.exe -m pytest
```

## Run the local release gate

Run the repo’s browser-first release validation in one command:

```powershell
.\scripts\validate-browser-release.ps1
```

This runs:

- `pytest`
- the real browser API + SSE corpus validation against `.\emails-for-testing`

Useful flags:

- `-KeepArtifacts` to leave `.local-browser-run\` reports on disk
- `-TimeoutSeconds 1200` to extend the corpus validation budget
- `-PythonPath .\.venv\Scripts\python.exe` to point at a different local interpreter

## Validate the canonical golden corpus

Validate the paired `.msg` + golden `.pdf` cases under `emails-for-testing`:

```powershell
.\.venv\Scripts\python.exe -m msg_to_pdf_dropzone.corpus_validator --cases-dir .\emails-for-testing
```

Or:

```powershell
msg-to-pdf-corpus-validate --cases-dir .\emails-for-testing
```

Optional flags:

- `--case 7` to validate one numbered case folder
- `--render-strategy fast` to compare fallback rendering behavior
- `--output-root .\.local-corpus-profiles` to control report location
- `--fail-on-warnings` to return a non-zero exit code when actionable warnings are present

## Validate the browser path against the canonical corpus

Run the real browser API plus SSE event stream against `emails-for-testing`, then compare the generated PDFs to the golden corpus:

```powershell
.\.venv\Scripts\python.exe -m msg_to_pdf_dropzone.browser_validation --cases-dir .\emails-for-testing
```

Or:

```powershell
msg-to-pdf-browser-validate --cases-dir .\emails-for-testing
```

By default this starts a temporary local browser server on a free localhost port, writes reports under `.local-browser-run\browser-validation-<timestamp>\`, and exits non-zero if the browser event flow or semantic PDF comparison fails.

The browser validation command starts the server in `fidelity` mode by default, so it exercises the real Outlook-first pipeline unless you explicitly override it with `--render-strategy fast`.

## Profile the local corpus

Run a local corpus profile against `emails-for-testing` and generate JSON + Markdown summaries:

```powershell
.\.venv\Scripts\python.exe -m msg_to_pdf_dropzone.corpus_profiler --runs 3 --emails-dir .\emails-for-testing
```

Reports are written under `.local-corpus-profiles\profile-<timestamp>\`.

## Render strategy (local tuning)

- Default (`fast`): HTML + Edge first, then ReportLab. This avoids Outlook-export artifacts that some messages produce.
- Optional (`fidelity`): Outlook MHTML + Edge first, then HTML + Edge, then ReportLab.

Set strategy for one shell session:

```powershell
$env:MSG_TO_PDF_RENDER_STRATEGY='fidelity'
```

## Task event log

The app can emit normalized task events to a JSONL file for external debugging tools.

Set the log path for one shell session:

```powershell
$env:MSG_TO_PDF_TASK_EVENT_LOG='C:\temp\msg-to-pdf-task-events.jsonl'
```

Then run the app normally:

```powershell
.\.venv\Scripts\python.exe -m msg_to_pdf_dropzone --no-browser
```

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

This is intended for local debugging, event-stream inspection, and regression analysis.

## Notes

- Maximum input files per batch is 25.
- Filenames are sanitized for Windows.
- Queue preview names are server-provided and intentionally use the latest sent date in each normalized email thread. See `docs/web-dropzone-contract.md`.
- If multiple outputs would have the same name, a numeric suffix is added.
- Outlook drag handling uses COM to export selected items to temporary `.msg` files when direct file paths are not provided.
- PDF generation now defaults to the HTML-to-PDF Edge path, then falls back to the built-in renderer. Set `MSG_TO_PDF_RENDER_STRATEGY=fidelity` to opt into the Outlook MHTML + Edge path first.

## Troubleshooting

- `Port already in use`
  Run the app on another port, for example `.\.venv\Scripts\python.exe -m msg_to_pdf_dropzone --port 8766`.

- `Browser validation falls back to edge_html more often than expected`
  This usually means the Outlook MHTML export path is unavailable for those messages on the current machine. Check that Outlook is installed, can open `.msg` files locally, and is not blocked by a first-run or profile prompt.

- `Outlook-first rendering is unavailable`
  The app defaults to `edge_html` already. If you explicitly set `MSG_TO_PDF_RENDER_STRATEGY=fidelity`, the Outlook-first path requires Windows Outlook automation plus Microsoft Edge.

- `Choose Output Folder does nothing`
  The native folder picker depends on local Windows UI access. Re-run the app from an interactive desktop session, not a headless or service context.

- `Browser validation or app startup fails after changing shells`
  Use the workspace interpreter directly: `.\.venv\Scripts\python.exe ...`. The local release script also assumes that virtualenv by default.
