@echo off
title Metasploitable3 Lab Build

echo =====================================
echo   Metasploitable3 Lab Build Launcher
echo =====================================
echo.

cd /d "%~dp0"

echo Checking for required tools...
echo.

echo Starting build...
echo.

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0new_build.ps1"

echo.
echo Build finished
pause