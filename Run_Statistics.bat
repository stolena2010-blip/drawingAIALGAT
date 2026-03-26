@echo off
setlocal

cd /d "%~dp0"

echo ========================================
echo   AI DRAW - Process Statistics
echo ========================================
echo.

if exist ".venv\Scripts\python.exe" (
    echo Starting process analysis...
    .venv\Scripts\python.exe process_analysis.py
) else (
    echo ERROR: .venv not found
    pause
)
