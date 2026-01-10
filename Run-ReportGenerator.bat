@echo off
title Report Card Generator

echo.
echo ============================================
echo   Report Card Generator - Starting...
echo ============================================
echo.

:: Check if Docker is running
docker info >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Docker Desktop is not running!
    echo.
    echo Please:
    echo   1. Open Docker Desktop from Start Menu
    echo   2. Wait for it to show "Running"
    echo   3. Run this file again
    echo.
    pause
    exit /b 1
)

echo [OK] Docker is running
echo.
echo Downloading and starting the app...
echo (This may take a few minutes the first time)
echo.

:: Stop any existing container
docker stop reportgen 2>nul
docker rm reportgen 2>nul

:: Pull latest image and run
docker run -d --name reportgen -p 8080:8080 ghcr.io/utkarsh-koppikar/reportgen:latest

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to start. Check Docker Desktop for errors.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   SUCCESS! App is running!
echo ============================================
echo.
echo   Opening browser...
echo.
echo   If browser doesn't open, go to:
echo   http://localhost:8080
echo.
echo   To stop: Close this window or run Stop-ReportGenerator.bat
echo ============================================

:: Wait a moment for container to start
timeout /t 3 /nobreak >nul

:: Open browser
start http://localhost:8080

echo.
echo Press any key to stop the app and exit...
pause >nul

:: Cleanup
docker stop reportgen >nul 2>&1
docker rm reportgen >nul 2>&1
echo App stopped. Goodbye!

