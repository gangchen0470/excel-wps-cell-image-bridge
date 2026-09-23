@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Repair-Update-Registration.ps1"
if errorlevel 1 echo Repair failed. Please keep this window and report the error.
pause
