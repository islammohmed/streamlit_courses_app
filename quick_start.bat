@echo off
chcp 65001 > nul
title نظام إدارة الدورات التدريبية
echo.
echo =================================
echo   نظام إدارة الدورات التدريبية
echo   Training Courses Management
echo =================================
echo.

REM Check for Python
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :run
)

python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :run
)

echo ❌ Python not found!
pause
exit /b 1

:run
echo ✓ Starting application...
echo 🌐 Opening at: http://localhost:8501
echo 🔄 Press Ctrl+C to stop
echo.

%PYTHON_CMD% -m streamlit run app.py

pause
