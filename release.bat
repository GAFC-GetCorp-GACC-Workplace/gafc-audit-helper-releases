@echo off
:: GAFC Audit Helper - Quick Release Script
:: Double-click this file to create a new release

setlocal enabledelayedexpansion

echo ============================================
echo   GAFC Audit Helper - Quick Release
echo ============================================
echo.

:: Check if in git bash or cmd
where bash >nul 2>&1
if %errorlevel% equ 0 (
    echo Running in Git Bash mode...
    bash -c "./release.sh"
    exit /b
)

:: Ask for version
set /p VERSION="Enter version number (e.g., 1.0.1): "
if "%VERSION%"=="" (
    echo Error: Version required!
    pause
    exit /b 1
)

:: Ask for release message
set /p MESSAGE="Enter release message (optional): "
if "%MESSAGE%"=="" (
    set MESSAGE=Release version %VERSION%
)

echo.
echo Creating release v%VERSION%...
echo.

:: Step 0: Build prod XLAM (always use gafc_audit_helper_new.xlam, then rename)
set OUT_PROD=gafc_audit_helper_new.xlam
set FINAL_XLAM=gafc_audit_helper.xlam

echo [0/5] Building XLAM (prod)...
python rebuild_xlam.py
if errorlevel 1 (
    echo Build failed. Fix errors and try again.
    pause
    exit /b 1
)
if not exist "%OUT_PROD%" (
    echo ERROR: %OUT_PROD% not found. Build did not produce prod output.
    pause
    exit /b 1
)
copy /Y "%OUT_PROD%" "%FINAL_XLAM%" >nul
del "%OUT_PROD%" >nul 2>&1
echo   OK -> %FINAL_XLAM%
echo.

:: Step 1: Add and commit changes
echo [1/5] Committing changes...
git add .
git commit -m "Release v%VERSION% - %MESSAGE%"

:: Step 2: Create tag
echo [2/5] Creating tag v%VERSION%...
git tag -a v%VERSION% -m "Release v%VERSION%"

:: Step 3: Push commits
echo [3/5] Pushing commits...
git push origin main

:: Step 4: Push tag
echo [4/5] Pushing tag (this triggers auto-release)...
git push origin v%VERSION%

echo [5/5] Done pushing

echo.
echo ============================================
echo   SUCCESS! Release v%VERSION% triggered!
echo ============================================
echo.
echo GitHub Actions is now running...
echo.
echo Check progress at:
echo https://github.com/muaroi2002/gafc-audit-helper/actions
echo.
echo Public release will be created at:
echo https://github.com/muaroi2002/gafc-audit-helper-releases/releases
echo.

pause
