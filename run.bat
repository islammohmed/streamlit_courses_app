@echo off
chcp 65001 > nul
title نظام إدارة الدورات التدريبية - Training Courses Management System
echo =================================
echo   نظام إدارة الدورات التدريبية
echo   Training Courses Management
echo =================================
echo.

REM Try to run the PowerShell script first
echo Attempting to run PowerShell setup script...
powershell -ExecutionPolicy Bypass -File "setup_and_run.ps1"

if %errorlevel% neq 0 (
    echo.
    echo PowerShell script failed. Trying basic batch approach...
    echo.
    
    REM Fallback to basic batch commands
    py --version >nul 2>&1
    if %errorlevel% neq 0 (
        python --version >nul 2>&1
        if %errorlevel% neq 0 (
            echo Python is not installed or not in PATH.
            echo Please install Python from https://python.org
            echo Make sure to check "Add Python to PATH" during installation.
            pause
            exit /b 1
        )
        set PYTHON_CMD=python
    ) else (
        set PYTHON_CMD=py
    )
    
    echo Installing requirements...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo Failed to install requirements.
        pause
        exit /b 1
    )
    
    echo Starting the application...
    echo The application will open in your web browser at http://localhost:8501
    echo Press Ctrl+C to stop the application
    echo.
    %PYTHON_CMD% -m streamlit run app.py
)

pause
