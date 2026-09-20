@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-ExcelAddin.ps1"
if errorlevel 1 pause
exit /b %errorlevel%
