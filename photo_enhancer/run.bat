@echo off
python main.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Application exited with error.
    pause
)
