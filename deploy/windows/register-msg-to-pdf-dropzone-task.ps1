param(
    [string]$TaskName = "msg-to-pdf-dropzone Web",
    [string]$AppRoot = "C:\Program Files\msg-to-pdf-dropzone",
    [string]$EnvFile = "C:\ProgramData\msg-to-pdf-dropzone\config\app.env",
    [string]$RunnerScript = (Join-Path $PSScriptRoot "run-msg-to-pdf-dropzone.ps1"),
    [string]$TaskUser = "",
    [string]$TaskPassword = ""
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $RunnerScript)) {
    throw "Runner script not found: $RunnerScript"
}

$taskAction = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$RunnerScript`" -AppRoot `"$AppRoot`" -EnvFile `"$EnvFile`""
$taskTrigger = New-ScheduledTaskTrigger -AtStartup
$taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable

if ([string]::IsNullOrWhiteSpace($TaskUser)) {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $taskAction `
        -Trigger $taskTrigger `
        -Settings $taskSettings `
        -Description "Starts msg-to-pdf-dropzone on host startup via run-msg-to-pdf-dropzone.ps1." `
        -RunLevel Highest `
        -Force | Out-Null
} else {
    if ([string]::IsNullOrWhiteSpace($TaskPassword)) {
        throw "TaskPassword is required when TaskUser is provided."
    }
    Register-ScheduledTask `
        -TaskName $TaskName `
        -User $TaskUser `
        -Password $TaskPassword `
        -Action $taskAction `
        -Trigger $taskTrigger `
        -Settings $taskSettings `
        -Description "Starts msg-to-pdf-dropzone on host startup via run-msg-to-pdf-dropzone.ps1." `
        -RunLevel Highest `
        -Force | Out-Null
}

Write-Host "Scheduled task registered: $TaskName"
