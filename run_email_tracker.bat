@echo off
cd /d "%~dp0"
py email_tracker.py
if %errorlevel% neq 0 (
    echo.
    echo ERROR: tracker failed. See message above.
    pause
)
