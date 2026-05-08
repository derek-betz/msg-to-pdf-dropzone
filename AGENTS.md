# Codex Safety Guardrails (Workspace Policy)

These instructions apply to this repository only.

## Scope
- Operate only inside `C:\AI\msg-to-pdf-dropzone` unless the user explicitly asks for cross-repo or system-wide work.

## Destructive Actions
- Do not run destructive commands without explicit user confirmation in the current thread.
- Examples: `git reset --hard`, `git clean -fd`, mass delete operations, registry edits, system-level service changes.

## Git Safety
- Never force-push unless explicitly requested.
- Do not amend or rewrite history unless explicitly requested.
- Prefer showing planned git actions before running multi-step history operations.

## Python/Dependency Safety
- Prefer the workspace virtualenv at `.\.venv\Scripts\python.exe`.
- Avoid global Python/package installs unless explicitly requested.

## Network/External Effects
- Ask before commands that publish or mutate external systems (e.g., `git push`, package publish, cloud deploy).

## Transparency
- Before substantial edits, state intended files and command plan briefly.
- After running important commands, summarize key results.

## WEB-SVR03 Live Deployment Notes
- Do not assume remote access is required. Codex sessions for this repo may already be running on `WEB-SVR03`; check `hostname` and `$env:COMPUTERNAME` before trying WinRM/admin shares.
- The live durable app root is `C:\Program Files\msg-to-pdf-dropzone`; older notes may mention `C:\incoming\msg-to-pdf-dropzone`, but live assets have been served from `C:\Program Files\msg-to-pdf-dropzone`.
- The live scheduled task is `msg-to-pdf-dropzone Web`, running as `SYSTEM`, with `C:\ProgramData\msg-to-pdf-dropzone\config\app.env`.
- For a live update, first validate the checkout locally, then take a timestamped backup under `C:\incoming\backups\`, copy the checkout into `C:\Program Files\msg-to-pdf-dropzone` while preserving `.venv`, reinstall with the live venv, and restart the scheduled task.
- After deployment, verify `https://emailpdf.hanson-inc.com/api/health`, `https://emailpdf.hanson-inc.com/api/settings`, and compare live static assets against the checkout. Use a no-proxy HTTPS check if proxy environment variables point to `127.0.0.1:9`.
