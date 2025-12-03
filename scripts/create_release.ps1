# Create Release for GAFC Audit Helper
# Usage: .\create_release.ps1 -Version "1.0.1" [-Message "Release notes"]

param(
    [Parameter(Mandatory=$true)]
    [string]$Version,

    [Parameter(Mandatory=$false)]
    [string]$Message = "Release version $Version"
)

$ErrorActionPreference = "Stop"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  GAFC Audit Helper Release Creator" -ForegroundColor Cyan
Write-Host "  Version: $Version" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoDir = Split-Path -Parent $scriptDir
$xlam = Join-Path $repoDir "gafc_audit_helper.xlam"
$manifestPath = Join-Path $repoDir "releases\audit_tool.json"

# Step 1: Verify XLAM exists
Write-Host "`n[1/7] Checking XLAM file..." -ForegroundColor Yellow
if (-not (Test-Path $xlam)) {
    Write-Error "XLAM file not found: $xlam`nPlease build the XLAM in Excel first!"
}
Write-Host "  ✓ Found: gafc_audit_helper.xlam" -ForegroundColor Green

# Step 2: Calculate SHA256
Write-Host "`n[2/7] Calculating SHA256 hash..." -ForegroundColor Yellow
$sha256 = (Get-FileHash $xlam -Algorithm SHA256).Hash.ToLower()
Write-Host "  ✓ SHA256: $sha256" -ForegroundColor Green

# Step 3: Get file size
$fileSize = (Get-Item $xlam).Length
Write-Host "  ✓ File size: $([math]::Round($fileSize/1KB, 2)) KB" -ForegroundColor Green

# Step 4: Update manifest
Write-Host "`n[3/7] Updating manifest..." -ForegroundColor Yellow
$date = Get-Date -Format "yyyy-MM-dd"

# Read existing manifest to preserve changelog
$existingManifest = $null
if (Test-Path $manifestPath) {
    $existingManifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
}

$manifest = @{
    latest = $Version
    download_url = "https://github.com/muaroi2002/gafc-audit-helper-releases/releases/download/v$Version/gafc_audit_helper.xlam"
    sha256 = $sha256
    release_date = $date
    release_notes = $Message
    changelog = if ($existingManifest.changelog) { $existingManifest.changelog } else {
        @(
            "License validation with online/offline grace periods",
            "State file encryption and checksum validation",
            "Clock tampering detection",
            "Hardware fingerprinting with multi-tier fallback",
            "Dynamic validation intervals based on expiry proximity"
        )
    }
    min_excel_version = "2016"
    file_size_bytes = $fileSize
}

$manifest | ConvertTo-Json -Depth 10 | Set-Content $manifestPath -Encoding UTF8
Write-Host "  ✓ Manifest updated: releases/audit_tool.json" -ForegroundColor Green

# Step 5: Git commit
Write-Host "`n[4/7] Committing changes to git..." -ForegroundColor Yellow
Push-Location $repoDir
try {
    git add gafc_audit_helper.xlam
    git add releases/audit_tool.json
    $commitMsg = "Release v$Version - $Message"
    git commit -m $commitMsg
    Write-Host "  ✓ Changes committed" -ForegroundColor Green
} catch {
    Write-Host "  ⚠ Warning: Git commit failed (maybe no changes?)" -ForegroundColor Yellow
}
Pop-Location

# Step 6: Create and push tag
Write-Host "`n[5/7] Creating git tag..." -ForegroundColor Yellow
Push-Location $repoDir
try {
    $tagExists = git tag -l "v$Version"
    if ($tagExists) {
        Write-Host "  ⚠ Tag v$Version already exists. Delete it? (Y/N)" -ForegroundColor Yellow
        $response = Read-Host
        if ($response -eq "Y" -or $response -eq "y") {
            git tag -d "v$Version"
            git push origin ":refs/tags/v$Version" 2>$null
            Write-Host "  ✓ Old tag deleted" -ForegroundColor Green
        } else {
            Write-Error "Cannot proceed with existing tag"
        }
    }

    git tag -a "v$Version" -m "Release v$Version"
    Write-Host "  ✓ Tag created: v$Version" -ForegroundColor Green
} catch {
    Write-Error "Failed to create tag: $_"
}
Pop-Location

# Step 7: Push to GitHub
Write-Host "`n[6/7] Pushing to GitHub..." -ForegroundColor Yellow
$pushNow = Read-Host "Push to GitHub now? (Y/N)"
if ($pushNow -eq "Y" -or $pushNow -eq "y") {
    Push-Location $repoDir
    try {
        git push origin main
        git push origin "v$Version"
        Write-Host "  ✓ Pushed to GitHub" -ForegroundColor Green
    } catch {
        Write-Error "Failed to push: $_"
    }
    Pop-Location
} else {
    Write-Host "  ⚠ Skipped push. Run manually:" -ForegroundColor Yellow
    Write-Host "    cd $repoDir" -ForegroundColor Gray
    Write-Host "    git push origin main" -ForegroundColor Gray
    Write-Host "    git push origin v$Version" -ForegroundColor Gray
}

# Step 8: Create GitHub Release (optional)
Write-Host "`n[7/7] Creating GitHub Release..." -ForegroundColor Yellow
Write-Host "  Options:" -ForegroundColor Cyan
Write-Host "    1. Create release via GitHub Web UI" -ForegroundColor Gray
Write-Host "    2. Use GitHub CLI (gh release create)" -ForegroundColor Gray
Write-Host "    3. Skip for now" -ForegroundColor Gray

$choice = Read-Host "Choose option (1-3)"

switch ($choice) {
    "1" {
        $repoUrl = git config --get remote.origin.url
        $repoUrl = $repoUrl -replace '\.git$', ''
        $repoUrl = $repoUrl -replace 'git@github.com:', 'https://github.com/'
        $releaseUrl = "$repoUrl/releases/new?tag=v$Version"
        Start-Process $releaseUrl
        Write-Host "  ✓ Opening GitHub in browser..." -ForegroundColor Green
    }
    "2" {
        if (Get-Command gh -ErrorAction SilentlyContinue) {
            Push-Location $repoDir
            try {
                gh release create "v$Version" `
                    gafc_audit_helper.xlam `
                    --title "Release v$Version" `
                    --notes "$Message`n`nSHA256: $sha256"
                Write-Host "  ✓ GitHub release created via CLI" -ForegroundColor Green
            } catch {
                Write-Error "Failed to create release: $_"
            }
            Pop-Location
        } else {
            Write-Host "  ✗ GitHub CLI not installed. Install: https://cli.github.com/" -ForegroundColor Red
        }
    }
    "3" {
        Write-Host "  ⚠ Skipped. Create manually on GitHub." -ForegroundColor Yellow
    }
}

# Summary
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  ✅ RELEASE v$Version READY!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  File: gafc_audit_helper.xlam" -ForegroundColor White
Write-Host "  SHA256: $sha256" -ForegroundColor White
Write-Host "  Manifest: Updated" -ForegroundColor White
Write-Host "  Tag: v$Version" -ForegroundColor White
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "  1. Verify release on GitHub" -ForegroundColor Gray
Write-Host "  2. Test auto-update: .\scripts\update_audit_helper.ps1" -ForegroundColor Gray
Write-Host "  3. Deploy to users" -ForegroundColor Gray

exit 0
