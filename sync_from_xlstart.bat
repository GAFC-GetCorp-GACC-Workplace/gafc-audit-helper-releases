@echo off
:: Sync XLAM file from XLSTART to Git folder before release
:: This ensures you're releasing the latest tested code

setlocal enabledelayedexpansion

echo ============================================
echo   Sync Code from XLSTART to Git
echo ============================================
echo.

set XLSTART=%APPDATA%\Microsoft\Excel\XLSTART
set GIT_FOLDER=%~dp0
set XLAM_NAME=gafc_audit_helper.xlam

set SOURCE=%XLSTART%\%XLAM_NAME%
set TARGET=%GIT_FOLDER%%XLAM_NAME%

echo Source: %SOURCE%
echo Target: %TARGET%
echo.

:: Check if source exists
if not exist "%SOURCE%" (
    echo [ERROR] XLAM file not found in XLSTART!
    echo Please install the add-in first.
    pause
    exit /b 1
)

:: Check if target exists
if exist "%TARGET%" (
    echo [WARNING] Target file exists and will be overwritten.
    choice /C YN /M "Continue? (Y/N)"
    if errorlevel 2 exit /b 0
)

:: Copy file
echo Copying...
copy /Y "%SOURCE%" "%TARGET%"

if errorlevel 1 (
    echo [ERROR] Copy failed!
    pause
    exit /b 1
)

echo.
echo ============================================
echo   SUCCESS! Code synced from XLSTART
echo ============================================
echo.
echo Next steps:
echo   1. Review changes: git diff gafc_audit_helper.xlam
echo   2. Commit: git add gafc_audit_helper.xlam
echo   3. Release: release.bat
echo.

pause
