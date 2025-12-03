# Remove Auto-Update for GAFC Audit Helper
# This script removes the scheduled task for auto-updates

$ErrorActionPreference = "Stop"

$TaskName = "GAFC Audit Helper Auto Update"

Write-Host "Removing auto-update for GAFC Audit Helper..." -ForegroundColor Cyan

$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Auto-update has been removed successfully!" -ForegroundColor Green
} else {
    Write-Host "No auto-update task found." -ForegroundColor Yellow
}

exit 0
