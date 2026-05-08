param(
    [string]$SourceRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path,
    [string]$AppRoot = "C:\Program Files\msg-to-pdf-dropzone",
    [string]$EnvFile = "C:\ProgramData\msg-to-pdf-dropzone\config\app.env",
    [string]$TaskName = "msg-to-pdf-dropzone Web",
    [string]$BackupRoot = "C:\incoming\backups",
    [string]$BaseUrl = "https://emailpdf.hanson-inc.com",
    [switch]$SkipTests,
    [switch]$SkipRestart
)

$ErrorActionPreference = "Stop"

function Invoke-RobocopyChecked {
    param(
        [string]$From,
        [string]$To,
        [string[]]$Options
    )

    & robocopy.exe $From $To @Options
    $code = $LASTEXITCODE
    if ($code -gt 7) {
        throw "robocopy failed with exit code $code from '$From' to '$To'."
    }
    Write-Host "robocopy exit code $code from '$From' to '$To'."
}

function Invoke-NativeChecked {
    param(
        [string]$Description,
        [scriptblock]$Script
    )

    Write-Host $Description
    & $Script
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE."
    }
}

function Get-GitValue {
    param([string[]]$GitArgs)

    try {
        $value = & git -C $SourceRoot @GitArgs 2>$null
        if ($LASTEXITCODE -eq 0) {
            return ($value | Select-Object -First 1).Trim()
        }
    } catch {
        return ""
    }
    return ""
}

function Invoke-LiveValidation {
    param(
        [string]$PythonExe,
        [string]$Url,
        [string]$Root,
        [string]$ExpectedRevision
    )

    $validationScript = @'
import hashlib
import json
import pathlib
import ssl
import sys
import urllib.request

base_url = sys.argv[1].rstrip("/")
app_root = pathlib.Path(sys.argv[2])
expected_revision = sys.argv[3]
opener = urllib.request.build_opener(
    urllib.request.ProxyHandler({}),
    urllib.request.HTTPSHandler(context=ssl._create_unverified_context()),
)

def fetch(path):
    return opener.open(base_url + path, timeout=20).read()

health = json.loads(fetch("/api/health").decode("utf-8"))
settings = json.loads(fetch("/api/settings").decode("utf-8"))
version = json.loads(fetch("/api/version").decode("utf-8"))
if health.get("ok") is not True:
    raise SystemExit(f"Health did not report ok=true: {health}")
if settings.get("serverMode") is not True:
    raise SystemExit(f"Settings did not report hosted serverMode=true: {settings}")
if not version.get("sourceRevision"):
    raise SystemExit(f"Version payload did not include sourceRevision: {version}")
if expected_revision != "unknown" and version.get("sourceRevision") != expected_revision:
    raise SystemExit(
        f"Version payload revision did not match deployment: "
        f"expected={expected_revision} actual={version.get('sourceRevision')}"
    )

pairs = [
    ("index.html", "/", app_root / "src" / "msg_to_pdf_dropzone" / "web_ui" / "index.html"),
    ("app.css", "/static/app.css?v=filename-panel-1", app_root / "src" / "msg_to_pdf_dropzone" / "web_ui" / "app.css"),
    ("app.js", "/static/app.js?v=result-row-actions-2", app_root / "src" / "msg_to_pdf_dropzone" / "web_ui" / "app.js"),
    ("dropzone_controller.js", "/static/dropzone_controller.js", app_root / "src" / "msg_to_pdf_dropzone" / "web_ui" / "dropzone_controller.js"),
]
for label, path, local_path in pairs:
    remote = fetch(path)
    local = local_path.read_bytes()
    if remote != local:
        raise SystemExit(
            f"{label} did not match live bytes: "
            f"remote={hashlib.sha256(remote).hexdigest()[:12]} "
            f"local={hashlib.sha256(local).hexdigest()[:12]}"
        )

print(json.dumps({"health": health, "settings": settings, "version": version}, indent=2))
'@

    & $PythonExe -c $validationScript $Url $Root $ExpectedRevision
    if ($LASTEXITCODE -ne 0) {
        throw "Live validation failed with exit code $LASTEXITCODE."
    }
}

$SourceRoot = (Resolve-Path -LiteralPath $SourceRoot).Path
if (-not (Test-Path -LiteralPath $AppRoot)) {
    throw "AppRoot not found: $AppRoot"
}
if (-not (Test-Path -LiteralPath $EnvFile)) {
    throw "EnvFile not found: $EnvFile"
}

$sourcePython = Join-Path $SourceRoot ".venv\Scripts\python.exe"
$livePython = Join-Path $AppRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $livePython)) {
    throw "Live Python executable not found: $livePython"
}

$revision = Get-GitValue @("rev-parse", "HEAD")
if (-not $revision) {
    $revision = "unknown"
}
$branch = Get-GitValue @("branch", "--show-current")
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = Join-Path $BackupRoot "msg-to-pdf-dropzone-before-$($revision.Substring(0, [Math]::Min(12, $revision.Length)))-$stamp"

Write-Host "Deploying msg-to-pdf-dropzone to WEB-SVR03"
Write-Host "SourceRoot: $SourceRoot"
Write-Host "AppRoot: $AppRoot"
Write-Host "EnvFile: $EnvFile"
Write-Host "TaskName: $TaskName"
Write-Host "Revision: $revision"

if (-not $SkipTests) {
    $testPython = if (Test-Path -LiteralPath $sourcePython) { $sourcePython } else { $livePython }
    $previousPythonPath = $env:PYTHONPATH
    $env:PYTHONPATH = Join-Path $SourceRoot "src"
    try {
        Invoke-NativeChecked "Running pytest release gate" {
            & $testPython -m pytest --basetemp (Join-Path $SourceRoot ".pytest-tmp")
        }
    } finally {
        $env:PYTHONPATH = $previousPythonPath
    }
}

New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null
Invoke-RobocopyChecked -From $AppRoot -To $backup -Options @(
    "/E",
    "/XD", ".venv", ".git", ".pytest_cache", "__pycache__", ".local-browser-run",
    "/XF", "*.pyc", "*.pyo"
)

Invoke-RobocopyChecked -From $SourceRoot -To $AppRoot -Options @(
    "/E",
    "/XD", ".venv", ".git", ".pytest_cache", "__pycache__", ".local-browser-run", ".ruff_cache", ".mypy_cache",
    "/XF", "*.pyc", "*.pyo"
)

$releasePath = Join-Path $AppRoot "src\msg_to_pdf_dropzone\_release.json"
$releasePayload = [ordered]@{
    appName = "msg-to-pdf-dropzone"
    sourceRevision = $revision
    sourceBranch = $branch
    deployedAt = (Get-Date).ToUniversalTime().ToString("o")
    deployedBy = "$env:USERDOMAIN\$env:USERNAME"
    sourceRoot = $SourceRoot
}
$releasePayload | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $releasePath -Encoding UTF8

Set-Location -LiteralPath $AppRoot
Invoke-NativeChecked "Installing package into live virtualenv" {
    & $livePython -m pip install --no-deps .
}

if (-not $SkipRestart) {
    Write-Host "Restarting scheduled task: $TaskName"
    & schtasks.exe /End /TN $TaskName
    $endCode = $LASTEXITCODE
    if ($endCode -ne 0 -and $endCode -ne 128) {
        throw "schtasks /End failed with exit code $endCode."
    }
    Start-Sleep -Seconds 3
    Invoke-NativeChecked "Starting scheduled task: $TaskName" {
        & schtasks.exe /Run /TN $TaskName
    }
    Start-Sleep -Seconds 5
}

Invoke-LiveValidation -PythonExe $livePython -Url $BaseUrl -Root $AppRoot -ExpectedRevision $revision

Write-Host "Deployment complete."
Write-Host "Backup: $backup"
