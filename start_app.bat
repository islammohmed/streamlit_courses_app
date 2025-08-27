@echo off
chcp 65001 > nul
title نظام إدارة الدورات التدريبية
cls

REM Change to the script directory
cd /d "%~dp0"

echo.
echo =================================
echo   نظام إدارة الدورات التدريبية
echo   Training Courses Management
echo =================================
echo.
echo Starting application...
echo.
echo The app will open at: http://localhost:8501
echo Press Ctrl+C to stop the application
echo.

echo "" | py -m streamlit run app.py --server.headless true

pause
