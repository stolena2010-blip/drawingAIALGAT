@echo off
REM Quick update shortcut — double-click to update DrawingAI on server
REM Runs the PowerShell update script
echo.
echo DrawingAI Pro - Quick Update
echo ============================
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0update.ps1"
echo.
pause
