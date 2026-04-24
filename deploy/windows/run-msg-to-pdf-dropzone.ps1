param(
    [string]$AppRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path,
    [string]$EnvFile = "C:\ProgramData\msg-to-pdf-dropzone\config\app.env",
    [string]$BindHost = "127.0.0.1",
    [int]$Port = 8765,
    [string]$LogPath = "C:\ProgramData\msg-to-pdf-dropzone\logs\msg-to-pdf-dropzone.log"
)

$ErrorActionPreference = "Stop"

function Import-EnvFile {
    param([string]$PathValue)

    if (-not (Test-Path -LiteralPath $PathValue)) {
        Write-Host "Env file not found, continuing with existing environment: ${PathValue}"
        return
    }

    Get-Content -LiteralPath $PathValue | ForEach-Object {
        $line = $_.Trim()
        if (-not $line -or $line.StartsWith("#")) {
            return
        }
        $pair = $line -split "=", 2
        if ($pair.Count -ne 2) {
            throw "Invalid env line in ${PathValue}: $line"
        }
        [Environment]::SetEnvironmentVariable($pair[0].Trim(), $pair[1].Trim(), "Process")
    }
}

function Set-DefaultEnv {
    param(
        [string]$Name,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($Name, "Process"))) {
        [Environment]::SetEnvironmentVariable($Name, $Value, "Process")
    }
}

$pythonExe = Join-Path $AppRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python executable not found: $pythonExe"
}

Import-EnvFile -PathValue $EnvFile

$stagingRoot = "C:\ProgramData\msg-to-pdf-dropzone\staging"
$outputRoot = "C:\ProgramData\msg-to-pdf-dropzone\outputs\pdf"
$logsRoot = Split-Path -Parent $LogPath

foreach ($pathValue in @($stagingRoot, $outputRoot, $logsRoot)) {
    if (-not (Test-Path -LiteralPath $pathValue)) {
        New-Item -ItemType Directory -Path $pathValue -Force | Out-Null
    }
}

Set-DefaultEnv -Name "APP_ENV" -Value "production"
Set-DefaultEnv -Name "MSG_TO_PDF_HOST" -Value $BindHost
Set-DefaultEnv -Name "MSG_TO_PDF_PORT" -Value "$Port"
Set-DefaultEnv -Name "MSG_TO_PDF_SERVER_MODE" -Value "1"
Set-DefaultEnv -Name "MSG_TO_PDF_STAGING_DIR" -Value $stagingRoot
Set-DefaultEnv -Name "MSG_TO_PDF_OUTPUT_DIR" -Value $outputRoot
Set-DefaultEnv -Name "MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT" -Value "1"
Set-DefaultEnv -Name "MSG_TO_PDF_DISABLE_OUTPUT_PICKER" -Value "1"
Set-DefaultEnv -Name "MSG_TO_PDF_RENDER_STRATEGY" -Value "fast"

$effectiveHost = [Environment]::GetEnvironmentVariable("MSG_TO_PDF_HOST", "Process")
$effectivePort = [Environment]::GetEnvironmentVariable("MSG_TO_PDF_PORT", "Process")

Write-Host "Starting msg-to-pdf-dropzone"
Write-Host "AppRoot: $AppRoot"
Write-Host "EnvFile: $EnvFile"
Write-Host "Host: $effectiveHost"
Write-Host "Port: $effectivePort"
Write-Host "LogPath: $LogPath"
Write-Host ""

if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$previousErrorActionPreference = $ErrorActionPreference
try {
    $ErrorActionPreference = "Continue"
    & $pythonExe -m msg_to_pdf_dropzone.web_server --host $effectiveHost --port $effectivePort --no-browser 2>&1 | Tee-Object -FilePath $LogPath -Append
} finally {
    $ErrorActionPreference = $previousErrorActionPreference
}

exit $LASTEXITCODE
