@echo off
title DrawingAI Pro - Web UI
cd /d "%~dp0"

REM Verify we're in the right directory
if not exist ".venv\Scripts\python.exe" (
    echo ERROR: .venv not found in %CD%
    echo Run: python -m venv .venv
    pause
    exit /b 1
)

REM Kill any existing process on port 8501
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":8501 " ^| findstr "LISTENING"') do (
    if not "%%a"=="0" (
        taskkill /F /PID %%a >nul 2>&1
    )
)
timeout /t 2 /nobreak >nul

REM Open browser and start Streamlit (minimize this window)
start "" http://localhost:8501
echo Starting Streamlit server... (minimize this window)
echo.
.venv\Scripts\python.exe -m streamlit run streamlit_app/app.py --server.port 8501
pause
