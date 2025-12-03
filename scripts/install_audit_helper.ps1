$ErrorActionPreference = "Stop"

# Paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoDir   = Split-Path -Parent $scriptDir  # Go up one level to root
$xlStart   = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"

# Auto-detect XLAM file in repo root
$addinFiles = Get-ChildItem -Path $repoDir -Filter "*.xlam" -File
if ($addinFiles.Count -eq 0) {
    Write-Error "Cannot find any .xlam file in: $repoDir"
} elseif ($addinFiles.Count -gt 1) {
    Write-Host "Multiple .xlam files found. Please select:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $addinFiles.Count; $i++) {
        Write-Host "  [$($i+1)] $($addinFiles[$i].Name)"
    }
    $selection = Read-Host "Enter number (1-$($addinFiles.Count))"
    $addinFile = $addinFiles[[int]$selection - 1]
} else {
    $addinFile = $addinFiles[0]
}

$source = $addinFile.FullName
$addinName = $addinFile.Name
$target = Join-Path $xlStart $addinName

Write-Host "Installing: $addinName" -ForegroundColor Cyan

if (-not (Test-Path $xlStart)) {
    New-Item -ItemType Directory -Path $xlStart | Out-Null
}

# Copy add-in to XLSTART so Excel auto-loads it
Copy-Item -Path $source -Destination $xlStart -Force

# Remove Zone.Identifier if present
$zoneFile = $target + ":Zone.Identifier"
if (Test-Path $zoneFile) { Remove-Item $zoneFile -Force }

Write-Host "Installed successfully at: $xlStart" -ForegroundColor Green
Write-Host "Opening Excel to activate license..." -ForegroundColor Cyan

# Open Excel with the add-in loaded
Start-Process "excel.exe"

Write-Host "Done! Please enter your license key when prompted." -ForegroundColor Green
exit 0
