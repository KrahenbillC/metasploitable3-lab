@echo off
title Metasploitable3 Lab Build
echo ========================================
echo   Metasploitable3 Lab Build Launcher
echo ========================================
echo.

cd /d "%~dp0metasploitable3"

echo Checking for required tools...

where vagrant >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Vagrant is not installed.
    pause
    exit /b
)

where VBoxManage >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: VirtualBox is not installed.
    pause
    exit /b
)

echo.
echo Starting build...
powershell -ExecutionPolicy Bypass -File build.ps1

echo.
echo Build finished.
pause