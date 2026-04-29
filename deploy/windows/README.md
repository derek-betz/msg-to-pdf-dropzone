# Windows Deployment Baseline

Use this folder for the `WEB-SVR03` deployment shape that mirrors `CostEstimateGenerator`.

## Expected layout

- App code: `C:\Program Files\msg-to-pdf-dropzone`
- Virtualenv: `C:\Program Files\msg-to-pdf-dropzone\.venv`
- Config: `C:\ProgramData\msg-to-pdf-dropzone\config\app.env`
- Shared TLS dir recommended: `C:\ProgramData\SharedTls`
- Logs: `C:\ProgramData\msg-to-pdf-dropzone\logs\`
- Staging: `C:\ProgramData\msg-to-pdf-dropzone\staging\`
- Output root: `C:\ProgramData\msg-to-pdf-dropzone\outputs\pdf\`
- Feedback records: `C:\ProgramData\msg-to-pdf-dropzone\outputs\feedback\`
- Shared feedback email config recommended: `C:\ProgramData\SharedFeedback\feedback-email.json`

## Files

- `msg-to-pdf-dropzone.env.example`: baseline environment file for `WEB-SVR03`
- `web-svr03.app.env`: proposed runtime values for `WEB-SVR03`
- `run-msg-to-pdf-dropzone.ps1`: starts the hosted web app on loopback
- `register-msg-to-pdf-dropzone-task.ps1`: registers the supported startup path on `WEB-SVR03`
- `WEB-SVR03-CHECKLIST.md`: copy/paste bring-up and validation steps for the server

## Serving model

- Local bring-up: `uvicorn` on `127.0.0.1:8765`
- Final internal URL: direct TLS on `10.1.13.203:443` using the shared wildcard cert
- Final internal URL: `https://emailpdf.hanson-inc.com`
- Writable runtime data: `C:\ProgramData\msg-to-pdf-dropzone`
- Recommended shared wildcard cert path: `C:\ProgramData\SharedTls\hanson-inc-wildcard.crt` and `.key`
- Feedback email uses `MSG_TO_PDF_FEEDBACK_CONFIG_PATH` or the `MSG_TO_PDF_FEEDBACK_SMTP_*` environment variables.

## Hosted-mode notes

- The hosted deployment should set `MSG_TO_PDF_SERVER_MODE=1`.
- `MSG_TO_PDF_OUTPUT_DIR` points to the shared server-managed PDF output folder.
- Native folder selection and Outlook import are disabled in hosted mode by default.

## Bring-up order

1. Copy the repo to `C:\Program Files\msg-to-pdf-dropzone`.
2. Build the virtualenv and install dependencies.
3. Create the runtime directories under `C:\ProgramData\msg-to-pdf-dropzone`.
4. Copy `msg-to-pdf-dropzone.env.example` to the real config path for the localhost proof.
5. Run `run-msg-to-pdf-dropzone.ps1` manually on the server and validate `http://127.0.0.1:8765/api/health`.
6. Place the wildcard cert and key in the shared TLS directory or update the final env file to match IT's chosen shared path.
7. Replace the localhost proof config with `web-svr03.app.env` for final direct TLS on `443`.
8. Register or update the scheduled task.
9. Validate `https://emailpdf.hanson-inc.com`.
