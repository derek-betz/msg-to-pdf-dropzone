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
  'C:\ProgramData\msg-to-pdf-dropzone\logs', `
  'C:\ProgramData\msg-to-pdf-dropzone\staging', `
  'C:\ProgramData\msg-to-pdf-dropzone\outputs\pdf'
```

## 3. Build the virtualenv and install the app

```powershell
Set-Location 'C:\Program Files\msg-to-pdf-dropzone'
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install .
```

## 4. Install the WEB-SVR03 app env file

Copy the checked-in server values to the live config path:

```powershell
Copy-Item `
  'C:\Program Files\msg-to-pdf-dropzone\deploy\windows\web-svr03.app.env' `
  'C:\ProgramData\msg-to-pdf-dropzone\config\app.env' `
  -Force
```

The current values are:

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

## 6. Register startup

After the localhost proof passes:

```powershell
Set-Location 'C:\Program Files\msg-to-pdf-dropzone\deploy\windows'
powershell -NoProfile -ExecutionPolicy Bypass -File `
  '.\register-msg-to-pdf-dropzone-task.ps1' `
  -AppRoot 'C:\Program Files\msg-to-pdf-dropzone' `
  -EnvFile 'C:\ProgramData\msg-to-pdf-dropzone\config\app.env'
```

If IT wants a specific service account, rerun the same command with `-TaskUser` and `-TaskPassword`.

## 7. Final server-side validation

- Confirm the scheduled task starts cleanly after a reboot or manual run.
- Confirm the app log is being written to `C:\ProgramData\msg-to-pdf-dropzone\logs\msg-to-pdf-dropzone.log`.
- Confirm the IIS or reverse-proxy binding for `emailpdf.hanson-inc.com` points to `127.0.0.1:8765`.
- Validate `https://emailpdf.hanson-inc.com`.
