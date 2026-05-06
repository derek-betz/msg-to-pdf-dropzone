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

function Stop-ExistingListener {
    param(
        [string]$AppName,
        [string]$HostValue,
        [int]$PortValue
    )

    if ([string]::IsNullOrWhiteSpace($HostValue) -or $HostValue -in @("127.0.0.1", "localhost")) {
        Write-Host "Skipping listener cleanup for ${AppName}; host is ${HostValue}."
        return
    }

    $endpoint = "${HostValue}:${PortValue}"
    Write-Host "Checking for existing ${AppName} listener on ${endpoint}..."

    $pattern = "^\s*TCP\s+$([regex]::Escape($endpoint))\s+\S+\s+LISTENING\s+(\d+)\s*$"
    $listenerPids = netstat -ano -p tcp |
        ForEach-Object {
            if ($_ -match $pattern) {
                [int]$Matches[1]
            }
        } |
        Sort-Object -Unique

    foreach ($pidValue in $listenerPids) {
        if ($pidValue -le 0) {
            continue
        }

        Write-Host "Stopping existing ${AppName} listener process tree: PID ${pidValue}"
        & taskkill.exe /PID $pidValue /T /F
        if ($LASTEXITCODE -ne 0) {
            throw "Unable to stop existing ${AppName} listener PID ${pidValue} on ${endpoint}."
        }
    }

    Start-Sleep -Seconds 2

    $remaining = netstat -ano -p tcp |
        ForEach-Object {
            if ($_ -match $pattern) {
                [int]$Matches[1]
            }
        } |
        Sort-Object -Unique

    if ($remaining) {
        throw "Existing ${AppName} listener(s) still remain on ${endpoint}: $($remaining -join ', ')"
    }

    Write-Host "No existing ${AppName} listener remains on ${endpoint}."
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
$sourceRoot = Join-Path $AppRoot "src"
$existingPythonPath = [Environment]::GetEnvironmentVariable("PYTHONPATH", "Process")
$nextPythonPath = if ($existingPythonPath) { "$sourceRoot;$existingPythonPath" } else { $sourceRoot }
[Environment]::SetEnvironmentVariable("PYTHONPATH", $nextPythonPath, "Process")

Stop-ExistingListener -AppName "msg-to-pdf-dropzone" -HostValue $effectiveHost -PortValue ([int]$effectivePort)
Set-Location -LiteralPath $AppRoot

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
