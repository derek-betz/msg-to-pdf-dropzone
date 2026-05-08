# msg-to-pdf-dropzone on WEB-SVR03

Use this checklist to stand up `msg-to-pdf-dropzone` at `https://emailpdf.hanson-inc.com`.

## 1. Copy the app to the server

Place the repo at:

- `C:\Program Files\msg-to-pdf-dropzone`

## 2. Create the runtime folders

Run this in an elevated PowerShell session:

```powershell
New-Item -ItemType Directory -Force -Path `
  'C:\Program Files\msg-to-pdf-dropzone', `
  'C:\ProgramData\msg-to-pdf-dropzone\config', `
  'C:\ProgramData\SharedTls', `
  'C:\ProgramData\SharedFeedback', `
  'C:\ProgramData\msg-to-pdf-dropzone\logs', `
  'C:\ProgramData\msg-to-pdf-dropzone\staging', `
  'C:\ProgramData\msg-to-pdf-dropzone\outputs\pdf', `
  'C:\ProgramData\msg-to-pdf-dropzone\outputs\feedback'
```

## 3. Build the virtualenv and install the app

```powershell
Set-Location 'C:\Program Files\msg-to-pdf-dropzone'
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install .
```

## 4. Install the localhost proof app env file

Copy the baseline localhost values to the live config path:

```powershell
Copy-Item `
  'C:\Program Files\msg-to-pdf-dropzone\deploy\windows\msg-to-pdf-dropzone.env.example' `
  'C:\ProgramData\msg-to-pdf-dropzone\config\app.env' `
  -Force
```

The localhost proof values are:

```env
APP_ENV=production
MSG_TO_PDF_HOST=127.0.0.1
MSG_TO_PDF_PORT=8765
MSG_TO_PDF_SERVER_MODE=1
MSG_TO_PDF_STAGING_DIR=C:\ProgramData\msg-to-pdf-dropzone\staging
MSG_TO_PDF_OUTPUT_DIR=C:\ProgramData\msg-to-pdf-dropzone\outputs\pdf
MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT=1
MSG_TO_PDF_DISABLE_OUTPUT_PICKER=1
MSG_TO_PDF_RENDER_STRATEGY=fast
```

## 5. Manual localhost proof

```powershell
Set-Location 'C:\Program Files\msg-to-pdf-dropzone'
powershell -NoProfile -ExecutionPolicy Bypass -File `
  '.\deploy\windows\run-msg-to-pdf-dropzone.ps1' `
  -AppRoot 'C:\Program Files\msg-to-pdf-dropzone' `
  -EnvFile 'C:\ProgramData\msg-to-pdf-dropzone\config\app.env'
```

Leave that PowerShell window running and validate from a second window:

```powershell
Invoke-WebRequest 'http://127.0.0.1:8765/api/health'
Invoke-WebRequest 'http://127.0.0.1:8765/api/settings'
```

## 6. Switch to the final WEB-SVR03 direct-TLS config

After the localhost proof passes, replace the config with the final `443` values:

Make sure the shared wildcard cert and key are already present at:

- `C:\ProgramData\SharedTls\hanson-inc-wildcard.crt`
- `C:\ProgramData\SharedTls\hanson-inc-wildcard.key`

```powershell
Copy-Item `
  'C:\Program Files\msg-to-pdf-dropzone\deploy\windows\web-svr03.app.env' `
  'C:\ProgramData\msg-to-pdf-dropzone\config\app.env' `
  -Force
```

The final direct-TLS values are:

```env
APP_ENV=production
MSG_TO_PDF_HOST=10.1.13.203
MSG_TO_PDF_PORT=443
MSG_TO_PDF_SERVER_MODE=1
MSG_TO_PDF_STAGING_DIR=C:\ProgramData\msg-to-pdf-dropzone\staging
MSG_TO_PDF_OUTPUT_DIR=C:\ProgramData\msg-to-pdf-dropzone\outputs\pdf
MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT=1
MSG_TO_PDF_DISABLE_OUTPUT_PICKER=1
MSG_TO_PDF_RENDER_STRATEGY=fast
MSG_TO_PDF_TLS_CERTFILE=C:\ProgramData\SharedTls\hanson-inc-wildcard.crt
MSG_TO_PDF_TLS_KEYFILE=C:\ProgramData\SharedTls\hanson-inc-wildcard.key
MSG_TO_PDF_SERVER_NAMES=emailpdf.hanson-inc.com
# Optional, recommended when feedback email is enabled:
# MSG_TO_PDF_FEEDBACK_CONFIG_PATH=C:\ProgramData\SharedFeedback\feedback-email.json
```

## 7. Register startup

After the localhost proof passes:

```powershell
Set-Location 'C:\Program Files\msg-to-pdf-dropzone\deploy\windows'
powershell -NoProfile -ExecutionPolicy Bypass -File `
  '.\register-msg-to-pdf-dropzone-task.ps1' `
  -AppRoot 'C:\Program Files\msg-to-pdf-dropzone' `
  -EnvFile 'C:\ProgramData\msg-to-pdf-dropzone\config\app.env'
```

If IT wants a specific service account, rerun the same command with `-TaskUser` and `-TaskPassword`.

## 8. Final server-side validation

- Confirm the scheduled task starts cleanly after a reboot or manual run.
- Confirm the app log is being written to `C:\ProgramData\msg-to-pdf-dropzone\logs\msg-to-pdf-dropzone.log`.
- Validate `https://emailpdf.hanson-inc.com/api/health`, `https://emailpdf.hanson-inc.com/api/settings`, and `https://emailpdf.hanson-inc.com/api/version`.
- Submit a feedback test and confirm it is saved under `C:\ProgramData\msg-to-pdf-dropzone\outputs\feedback\`.

## Refresh an existing live install

If Codex is already running on `WEB-SVR03`, check `hostname` before trying remoting. The durable live root is:

- `C:\Program Files\msg-to-pdf-dropzone`

From a validated checkout, use the scripted refresh path:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  '.\deploy\windows\deploy-live-web-svr03.ps1' `
  -SourceRoot (Get-Location).Path
```

The script takes a timestamped backup under `C:\incoming\backups\`, preserves `.venv`, copies the checkout, writes release metadata, reinstalls the package into the live virtualenv, restarts `msg-to-pdf-dropzone Web`, and verifies the live health/settings/version endpoints plus static UI asset hashes.
