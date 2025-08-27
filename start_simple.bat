@echo off
chcp 65001 > nul
title نظام إدارة الدورات التدريبية - Training Courses Management System
echo.
echo =================================
echo   نظام إدارة الدورات التدريبية
echo   Training Courses Management
echo =================================
echo.

REM Try py first, then python
py --version >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ Python found using 'py' command
    set PYTHON_CMD=py
    goto :install
)

python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ Python found using 'python' command
    set PYTHON_CMD=python
    goto :install
)

echo ❌ Python is not installed or not in PATH.
echo.
echo Please install Python from https://python.org
echo Make sure to check "Add Python to PATH" during installation.
echo.
pause
exit /b 1

:install
echo.
echo Installing required packages...
%PYTHON_CMD% -m pip install streamlit pandas plotly openpyxl python-docx >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Failed to install some packages. Trying individual installation...
    %PYTHON_CMD% -m pip install streamlit
    %PYTHON_CMD% -m pip install pandas
    %PYTHON_CMD% -m pip install plotly
    %PYTHON_CMD% -m pip install openpyxl
    %PYTHON_CMD% -m pip install python-docx
)

echo ✓ Packages installed successfully!
echo.
echo Starting the application...
echo 🌐 The application will open at: http://localhost:8501
echo 📖 Use the web interface to upload your files or use the default ones
echo 🔄 Press Ctrl+C to stop the application
echo.

%PYTHON_CMD% -m streamlit run app.py

pause
