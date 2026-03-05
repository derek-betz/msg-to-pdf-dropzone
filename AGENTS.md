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
