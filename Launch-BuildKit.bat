@echo off
title Metasploitable3 Student Build Kit
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\new_build.ps1"
pause