[CmdletBinding()]
param(
    [string]$PythonPath = ".\.venv\Scripts\python.exe",
    [string]$CasesDir = ".\emails-for-testing",
    [string]$OutputRoot = ".\.local-browser-run",
    [string]$HostName = "127.0.0.1",
    [int]$Port = 8767,
    [int]$MaxCases = 4,
    [ValidateSet("date_subject", "subject", "sender_subject", "date_sender_subject")]
    [string]$FilenameStyle = "sender_subject"
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

$runStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$runDir = Join-Path (Join-Path $repoRoot $OutputRoot) "rendered-ui-workflow-$runStamp"
$outputDir = Join-Path $runDir "generated"
$serverLog = Join-Path $runDir "server.log"
$serverErrLog = Join-Path $runDir "server.err.log"
$manifestPath = Join-Path $runDir "manifest.json"

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$msgFiles = Get-ChildItem -Path $resolvedCasesDir -Recurse -Filter "*.msg" |
    Sort-Object FullName |
    Select-Object -First $MaxCases

if (-not $msgFiles) {
    throw "No .msg files were found under '$resolvedCasesDir'."
}

$oldOutputDir = $env:MSG_TO_PDF_OUTPUT_DIR
$oldDisablePicker = $env:MSG_TO_PDF_DISABLE_OUTPUT_PICKER
$oldRenderStrategy = $env:MSG_TO_PDF_RENDER_STRATEGY
$oldDisableOutlook = $env:MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT

try {
    $env:MSG_TO_PDF_OUTPUT_DIR = $outputDir
    $env:MSG_TO_PDF_DISABLE_OUTPUT_PICKER = "1"
    $env:MSG_TO_PDF_RENDER_STRATEGY = "fast"
    $env:MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT = "1"

    $process = Start-Process `
        -FilePath $resolvedPythonPath `
        -ArgumentList @("-m", "msg_to_pdf_dropzone", "--host", $HostName, "--port", [string]$Port, "--no-browser") `
        -WorkingDirectory $repoRoot `
        -WindowStyle Hidden `
        -RedirectStandardOutput $serverLog `
        -RedirectStandardError $serverErrLog `
        -PassThru
} finally {
    $env:MSG_TO_PDF_OUTPUT_DIR = $oldOutputDir
    $env:MSG_TO_PDF_DISABLE_OUTPUT_PICKER = $oldDisablePicker
    $env:MSG_TO_PDF_RENDER_STRATEGY = $oldRenderStrategy
    $env:MSG_TO_PDF_DISABLE_OUTLOOK_IMPORT = $oldDisableOutlook
}

$baseUrl = "http://${HostName}:$Port"
$deadline = (Get-Date).AddSeconds(30)
$healthy = $false
while ((Get-Date) -lt $deadline) {
    if ($process.HasExited) {
        throw "Rendered UI workflow server exited early with code $($process.ExitCode). See '$serverErrLog'."
    }
    try {
        $health = Invoke-RestMethod -Uri "$baseUrl/api/health" -TimeoutSec 2
        if ($health.ok -eq $true) {
            $healthy = $true
            break
        }
    } catch {
        Start-Sleep -Milliseconds 500
    }
}

if (-not $healthy) {
    Stop-Process -Id $process.Id -Force
    throw "Timed out waiting for rendered UI workflow server at '$baseUrl'."
}

$manifest = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    baseUrl = $baseUrl
    processId = $process.Id
    runDir = $runDir
    outputDir = $outputDir
    filenameStyle = $FilenameStyle
    files = @($msgFiles | ForEach-Object { $_.FullName })
    selectors = [ordered]@{
        dropzone = "[data-testid='dropzone']"
        fileInput = "[data-testid='file-input']"
        filenameStyleSelect = "[data-testid='filename-style-select']"
        filenameStyleExample = "[data-testid='filename-style-example']"
        queueList = "[data-testid='queue-list']"
        convertButton = "[data-testid='convert-button']"
        batchProgress = "[data-testid='batch-progress']"
        resultBanner = "[data-testid='result-banner']"
        resultHeadline = "[data-testid='result-headline']"
    }
    expectations = [ordered]@{
        queuedCount = $msgFiles.Count
        resultHeadlineContains = "$($msgFiles.Count) PDFs saved"
        outputNamesContainSenderPrefix = ($FilenameStyle -eq "sender_subject" -or $FilenameStyle -eq "date_sender_subject")
    }
}

$manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $manifestPath -Encoding UTF8

Write-Host "Rendered UI workflow server is ready: $baseUrl"
Write-Host "Manifest: $manifestPath"
Write-Host "Process ID: $($process.Id)"
Write-Host "Stop it with: Stop-Process -Id $($process.Id)"
