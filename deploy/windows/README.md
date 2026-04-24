# Windows Deployment Baseline

Use this folder for the `WEB-SVR03` deployment shape that mirrors `CostEstimateGenerator`.

## Expected layout

- App code: `C:\Program Files\msg-to-pdf-dropzone`
- Virtualenv: `C:\Program Files\msg-to-pdf-dropzone\.venv`
- Config: `C:\ProgramData\msg-to-pdf-dropzone\config\app.env`
- Logs: `C:\ProgramData\msg-to-pdf-dropzone\logs\`
- Staging: `C:\ProgramData\msg-to-pdf-dropzone\staging\`
- Output root: `C:\ProgramData\msg-to-pdf-dropzone\outputs\pdf\`

## Files

- `msg-to-pdf-dropzone.env.example`: baseline environment file for `WEB-SVR03`
- `web-svr03.app.env`: proposed runtime values for `WEB-SVR03`
- `run-msg-to-pdf-dropzone.ps1`: starts the hosted web app on loopback
- `register-msg-to-pdf-dropzone-task.ps1`: registers the supported startup path on `WEB-SVR03`
- `WEB-SVR03-CHECKLIST.md`: copy/paste bring-up and validation steps for the server

## Serving model

- App server: `uvicorn` on `127.0.0.1:8765`
- Final internal URL: `https://emailpdf.hanson-inc.com`
- Writable runtime data: `C:\ProgramData\msg-to-pdf-dropzone`

## Hosted-mode notes

- The hosted deployment should set `MSG_TO_PDF_SERVER_MODE=1`.
- `MSG_TO_PDF_OUTPUT_DIR` points to the shared server-managed PDF output folder.
- Native folder selection and Outlook import are disabled in hosted mode by default.

## Bring-up order

1. Copy the repo to `C:\Program Files\msg-to-pdf-dropzone`.
2. Build the virtualenv and install dependencies.
3. Create the runtime directories under `C:\ProgramData\msg-to-pdf-dropzone`.
4. Copy `web-svr03.app.env` to the real config path as `app.env`.
5. Run `run-msg-to-pdf-dropzone.ps1` manually on the server for the localhost proof.
6. Validate `http://127.0.0.1:8765/api/health`.
7. Add the web-server binding for `emailpdf.hanson-inc.com` on `WEB-SVR03`.
8. Register or update the scheduled task and validate the real URL.
