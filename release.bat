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

:: Step 1: Add and commit changes
echo [1/4] Committing changes...
git add .
git commit -m "Release v%VERSION% - %MESSAGE%"

:: Step 2: Create tag
echo [2/4] Creating tag v%VERSION%...
git tag -a v%VERSION% -m "Release v%VERSION%"

:: Step 3: Push commits
echo [3/4] Pushing commits...
git push origin main

:: Step 4: Push tag
echo [4/4] Pushing tag (this triggers auto-release)...
git push origin v%VERSION%

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
