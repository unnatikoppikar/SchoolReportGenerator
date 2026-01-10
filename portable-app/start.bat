@echo off
title Report Card Generator
echo.
echo ========================================
echo   Report Card Generator
echo   Starting server...
echo ========================================
echo.

cd /d "%~dp0"

REM Check for bundled Python
if exist "python\python.exe" (
    set PYTHON=python\python.exe
) else (
    set PYTHON=python
)

REM Check if Python exists
%PYTHON% --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo Please install Python 3.10+ or place Python in the 'python' folder.
    pause
    exit /b 1
)

REM Run the app
cd app
%PYTHON% main.py

pause

