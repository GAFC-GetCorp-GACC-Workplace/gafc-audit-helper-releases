# Sync Files to Public Releases Repo
# This script copies necessary files from private repo to public repo
# Usage: .\sync_to_public.ps1

$ErrorActionPreference = "Stop"

# Paths
$privateRepo = "E:\audit\GAFC_Audit_Helper_Release"
$publicRepo = "E:\audit\GAFC_Audit_Helper_Release_Public"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Sync to Public Releases Repo" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Verify paths exist
if (-not (Test-Path $privateRepo)) {
    Write-Error "Private repo not found: $privateRepo"
}

if (-not (Test-Path $publicRepo)) {
    Write-Error "Public repo not found: $publicRepo"
}

Write-Host "`n[1/5] Copying XLAM file..." -ForegroundColor Yellow
$xlam = Get-ChildItem -Path $privateRepo -Filter "gafc_audit_helper.xlam" -File
if (-not $xlam) {
    Write-Error "XLAM file not found in private repo"
}
Copy-Item $xlam.FullName -Destination $publicRepo -Force
Write-Host "  ✓ Copied: $($xlam.Name)" -ForegroundColor Green

Write-Host "`n[2/5] Copying user scripts..." -ForegroundColor Yellow
$userScripts = @(
    "install_audit_helper.ps1",
    "uninstall_audit_helper.ps1",
    "update_audit_helper.ps1",
    "setup_auto_update.ps1",
    "remove_auto_update.ps1"
)

foreach ($script in $userScripts) {
    $source = Join-Path $privateRepo "scripts\$script"
    $dest = Join-Path $publicRepo "scripts\$script"
    if (Test-Path $source) {
        Copy-Item $source -Destination $dest -Force
        Write-Host "  ✓ Copied: $script" -ForegroundColor Green
    } else {
        Write-Host "  ⚠ Skipped (not found): $script" -ForegroundColor Yellow
    }
}

Write-Host "`n[3/5] Copying manifest..." -ForegroundColor Yellow
$manifestSrc = Join-Path $privateRepo "releases\audit_tool.json"
$manifestDst = Join-Path $publicRepo "releases\audit_tool.json"
if (Test-Path $manifestSrc) {
    Copy-Item $manifestSrc -Destination $manifestDst -Force
    Write-Host "  ✓ Copied: audit_tool.json" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Manifest not found" -ForegroundColor Yellow
}

Write-Host "`n[4/5] Copying documentation..." -ForegroundColor Yellow
$docFiles = @(
    "AUTO_UPDATE_SETUP.md"
)

foreach ($doc in $docFiles) {
    $source = Join-Path $privateRepo $doc
    $dest = Join-Path $publicRepo $doc
    if (Test-Path $source) {
        Copy-Item $source -Destination $dest -Force
        Write-Host "  ✓ Copied: $doc" -ForegroundColor Green
    }
}

Write-Host "`n[5/5] Creating installer package..." -ForegroundColor Yellow
# Create a zip package for easy distribution
$zipPath = Join-Path $publicRepo "gafc_audit_helper_installer.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

$filesToZip = @(
    (Join-Path $publicRepo "gafc_audit_helper.xlam"),
    (Join-Path $publicRepo "scripts\install_audit_helper.ps1"),
    (Join-Path $publicRepo "scripts\setup_auto_update.ps1"),
    (Join-Path $publicRepo "scripts\uninstall_audit_helper.ps1"),
    (Join-Path $publicRepo "AUTO_UPDATE_SETUP.md")
)

# Create temp directory for zip structure
$tempZipDir = Join-Path $env:TEMP "gafc_installer_temp"
if (Test-Path $tempZipDir) {
    Remove-Item $tempZipDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempZipDir | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempZipDir "scripts") | Out-Null

# Copy files to temp structure
Copy-Item (Join-Path $publicRepo "gafc_audit_helper.xlam") -Destination $tempZipDir -Force
Copy-Item (Join-Path $publicRepo "scripts\install_audit_helper.ps1") -Destination (Join-Path $tempZipDir "scripts") -Force
Copy-Item (Join-Path $publicRepo "scripts\setup_auto_update.ps1") -Destination (Join-Path $tempZipDir "scripts") -Force
Copy-Item (Join-Path $publicRepo "scripts\uninstall_audit_helper.ps1") -Destination (Join-Path $tempZipDir "scripts") -Force
Copy-Item (Join-Path $publicRepo "AUTO_UPDATE_SETUP.md") -Destination $tempZipDir -Force -ErrorAction SilentlyContinue

# Create README for installer
$installerReadme = @"
# GAFC Audit Helper Installer

## Quick Start

1. Run PowerShell as Administrator
2. Navigate to this folder
3. Run: .\scripts\install_audit_helper.ps1

## Setup Auto-Update (Recommended)

After installation, run:
.\scripts\setup_auto_update.ps1

See AUTO_UPDATE_SETUP.md for detailed instructions.
"@
Set-Content -Path (Join-Path $tempZipDir "README.txt") -Value $installerReadme -Encoding UTF8

# Create zip
Compress-Archive -Path "$tempZipDir\*" -DestinationPath $zipPath -Force
Remove-Item $tempZipDir -Recurse -Force

Write-Host "  ✓ Created: gafc_audit_helper_installer.zip" -ForegroundColor Green

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  ✅ Sync Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "`nFiles synced to: $publicRepo" -ForegroundColor White
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "  1. cd $publicRepo" -ForegroundColor Gray
Write-Host "  2. git add ." -ForegroundColor Gray
Write-Host "  3. git commit -m 'Update files'" -ForegroundColor Gray
Write-Host "  4. git push" -ForegroundColor Gray

exit 0
