@echo off
chcp 65001 >nul
title Unified Media Organizer

echo ========================================
echo Media Organizer (Windows Launcher)
echo ========================================

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ==============================================================================
    echo [ERROR] Python is not installed or not in your system PATH.
    echo.
    echo Please follow these steps:
    echo 1. Download Python 3.9 or newer from: https://www.python.org/downloads/
    echo 2. Run the installer.
    echo 3. *** CRITICAL: Check the box "Add python.exe to PATH" at the bottom of the screen! ***
    echo 4. Click "Install Now".
    echo 5. After installation, close this window and try running this script again.
    echo ==============================================================================
    echo.
    pause
    exit /b 1
)
:: Clean up unused Mac/Linux and Docker files to save space
if exist install_and_run.sh del install_and_run.sh
if exist Dockerfile del Dockerfile
if exist docker-compose.yml del docker-compose.yml

:: Run the universal Python launcher
python run.py

echo.
pause
