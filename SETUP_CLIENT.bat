@echo off
echo ====================================
echo   Course Management System Setup
echo ====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please:
    echo 1. Go to https://python.org/downloads
    echo 2. Download and install Python
    echo 3. Make sure to check "Add Python to PATH" during installation
    echo 4. Restart your computer
    echo 5. Run this script again
    echo.
    pause
    exit /b 1
)

echo Python is installed - checking version...
python --version

echo.
echo Installing required libraries...
python -m pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install requirements
    echo Make sure you have internet connection
    echo.
    pause
    exit /b 1
)

echo.
echo ====================================
echo   Starting the application...
echo ====================================
echo.
echo The application will open in your web browser
echo Go to: http://localhost:8501
echo.
echo Press Ctrl+C to stop the application
echo.

python -m streamlit run app.py

pause
