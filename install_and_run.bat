@echo off
chcp 65001 >nul
title Unified Media Organizer

echo ========================================
echo Media Organizer (Windows Launcher)
echo ========================================

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in your system PATH.
    echo Please download and install Python 3.9 or newer from https://www.python.org/downloads/
    echo During installation, make sure to check the box "Add Python to PATH".
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
