# Setup Auto-Update for GAFC Audit Helper
# This script creates a Windows Scheduled Task to automatically check for updates every 12 hours

$ErrorActionPreference = "Stop"

# Configuration
$TaskName = "GAFC Audit Helper Auto Update"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$UpdateScript = Join-Path $ScriptDir "update_audit_helper.ps1"
$UpdateInterval = 12  # Hours between update checks

# Verify update script exists
if (-not (Test-Path $UpdateScript)) {
    Write-Error "Update script not found: $UpdateScript"
}

Write-Host "Setting up auto-update for GAFC Audit Helper..." -ForegroundColor Cyan
Write-Host "Update check interval: Every $UpdateInterval hours" -ForegroundColor Cyan

# Remove existing task if present
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing scheduled task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Create scheduled task action (run PowerShell hidden)
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$UpdateScript`""

# Create trigger: repeat every N hours, indefinitely
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Hours $UpdateInterval)

# Task settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

# Register the task (run as current user, no elevation needed)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Automatically checks for and installs GAFC Audit Helper updates when Excel is not running" | Out-Null

Write-Host "`nAuto-update has been configured successfully!" -ForegroundColor Green
Write-Host "`nTask Details:" -ForegroundColor Cyan
Write-Host "  - Name: $TaskName"
Write-Host "  - Interval: Every $UpdateInterval hours"
Write-Host "  - First check: In 5 minutes"
Write-Host "  - Log file: $env:TEMP\gafc_update.log"
Write-Host "`nThe updater will run silently in the background."
Write-Host "It will only update when Excel is closed."

# Offer to run first check now
$response = Read-Host "`nDo you want to run the first update check now? (Y/N)"
if ($response -eq "Y" -or $response -eq "y") {
    Write-Host "`nRunning update check..." -ForegroundColor Cyan
    & $UpdateScript
}

Write-Host "`nSetup complete!" -ForegroundColor Green
exit 0
