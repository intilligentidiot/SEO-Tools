@echo off
title SEO & Web Tools Suite
echo ========================================================
echo   Starting SEO & Web Tools Suite (Local Development)
echo ========================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python was not found in PATH!
    echo Please make sure Python 3.9+ is installed.
    pause
    exit /b
)

echo Installing / verifying dependencies...
python -m pip install -r api/requirements.txt

echo.
echo Starting server at http://localhost:5000 ...
start http://localhost:5000
python api/index.py

pause
