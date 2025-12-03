# Convert all VBA text modules to UTF-8 with BOM so Unicode (tiếng Việt) renders đúng dấu
# Usage:
#   pwsh .\convert_vba_utf8bom.ps1 -Path .\ -Recurse
#   pwsh .\convert_vba_utf8bom.ps1 -Path .\modLicenseAudit.bas.vba
# Options:
#   -Path            File or folder. Defaults to current folder.
#   -Recurse         Include subfolders (when Path is a folder).
#   -SourceEncoding  Codepage to read (Default, 1258, 1252...). Default = system ANSI.
#   -NoBackup        Skip creating .bak copies.
param(
    [string]$Path = ".",
    [switch]$Recurse = $false,
    [string]$SourceEncoding = "Default",
    [switch]$NoBackup = $false
)

function Get-EncodingObject {
    param([string]$Name)
    if ($Name -eq "Default") { return [System.Text.Encoding]::Default }
    return [System.Text.Encoding]::GetEncoding($Name)
}

$sourceEnc = Get-EncodingObject $SourceEncoding
$utf8Bom = New-Object System.Text.UTF8Encoding $true

# Collect files
if (Test-Path $Path -PathType Leaf) {
    $files = @(Get-Item $Path)
} elseif (Test-Path $Path -PathType Container) {
    $files = Get-ChildItem -Path $Path -File -Recurse:$Recurse | Where-Object { $_.Extension -in ".bas", ".cls", ".frm" }
} else {
    Write-Error "Path not found: $Path"
    exit 1
}

if (-not $files.Count) {
    Write-Host "No VBA text modules found."
    exit 0
}

foreach ($file in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $text = $sourceEnc.GetString($bytes)
    if (-not $NoBackup) {
        Copy-Item $file.FullName "$($file.FullName).bak" -Force
    }
    [System.IO.File]::WriteAllText($file.FullName, $text, $utf8Bom)
    Write-Host "Converted -> UTF-8 BOM:" $file.FullName
}

Write-Host "Done. Re-import these modules into VBA so tiếng Việt hiển thị đúng dấu."
