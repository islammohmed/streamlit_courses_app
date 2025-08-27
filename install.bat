@echo off
echo Installing Streamlit Training Courses Management System...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

echo Python found. Installing requirements...
echo.

REM Install requirements
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo Failed to install requirements.
    pause
    exit /b 1
)

echo.
echo Installation completed successfully!
echo.
echo To run the application:
echo 1. Open command prompt or PowerShell
echo 2. Navigate to this directory
echo 3. Run: streamlit run app.py
echo.
echo Sample data files are available in the sample_data folder.
echo.
pause
