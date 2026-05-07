[CmdletBinding()]
param(
    [string]$PythonPath = ".\.venv\Scripts\python.exe",
    [string]$CasesDir = ".\emails-for-testing",
    [string]$OutputRoot = ".\.local-browser-run",
    [double]$TimeoutSeconds = 900,
    [switch]$KeepArtifacts
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$resolvedPythonPath = Join-Path $repoRoot $PythonPath
if (-not (Test-Path $resolvedPythonPath)) {
    throw "Python executable not found at '$resolvedPythonPath'. Create the workspace virtualenv first."
}

$resolvedCasesDir = Join-Path $repoRoot $CasesDir
if (-not (Test-Path $resolvedCasesDir)) {
    throw "Cases directory not found at '$resolvedCasesDir'."
}

$resolvedOutputRoot = Join-Path $repoRoot $OutputRoot
$existingValidationDirs = @{}
if (Test-Path $resolvedOutputRoot) {
    Get-ChildItem -Path $resolvedOutputRoot -Directory -Filter "browser-validation-*" | ForEach-Object {
        $existingValidationDirs[$_.FullName] = $true
    }
}

Write-Host "Running unit and integration tests..."
& $resolvedPythonPath -m pytest
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "Running canonical browser validation..."
& $resolvedPythonPath -m msg_to_pdf_dropzone.browser_validation --cases-dir $resolvedCasesDir --output-root $resolvedOutputRoot --timeout-seconds $TimeoutSeconds
$browserValidationExitCode = $LASTEXITCODE
if ($browserValidationExitCode -ne 0) {
    Write-Error "Browser validation failed with exit code $browserValidationExitCode. Validation artifacts were kept at '$resolvedOutputRoot'."
    exit $browserValidationExitCode
}

if (-not $KeepArtifacts -and (Test-Path $resolvedOutputRoot)) {
    Get-ChildItem -Path $resolvedOutputRoot -Directory -Filter "browser-validation-*" | Where-Object {
        -not $existingValidationDirs.ContainsKey($_.FullName)
    } | ForEach-Object {
        Remove-Item $_.FullName -Recurse -Force
    }
}

Write-Host "Release validation passed."
exit 0
