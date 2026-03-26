@echo off
setlocal

cd /d "%~dp0"

echo ========================================
echo   DrawingAI Pro - GUI Launcher
echo ========================================
echo.

if exist "automation_main.py" (
    set "TARGET_SCRIPT=automation_main.py"
) else if exist "customer_extractor_gui.py" (
    set "TARGET_SCRIPT=customer_extractor_gui.py"
) else if exist "main.py" (
    set "TARGET_SCRIPT=main.py"
) else (
    echo ERROR: GUI script not found. Missing automation_main.py / customer_extractor_gui.py / main.py
    echo Current folder: %CD%
    pause
    exit /b 1
)

echo Target: %TARGET_SCRIPT%

if exist ".venv\Scripts\python.exe" (
    echo Using virtual environment Python...
    ".venv\Scripts\python.exe" "%TARGET_SCRIPT%"
    goto :after_run
)

where py >nul 2>&1
if %errorlevel%==0 (
    echo .venv not found. Using py launcher...
    py -3 "%TARGET_SCRIPT%"
    goto :after_run
)

where python >nul 2>&1
if %errorlevel%==0 (
    echo .venv not found. Using system python...
    python "%TARGET_SCRIPT%"
    goto :after_run
)

echo ERROR: Python was not found.
echo Install Python or create .venv in this project folder.
pause
exit /b 1

:after_run
if errorlevel 1 (
    echo.
    echo ERROR: Project failed to start.
    echo Check the error message above.
    pause
    exit /b 1
)

exit /b 0
