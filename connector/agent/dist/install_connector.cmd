@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_wizard.ps1"
exit /b %errorlevel%
