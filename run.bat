@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating local Python virtual environment...
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if errorlevel 1 (
        echo Failed to create virtual environment. Please install Python 3 first.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install dependencies. Please check network or Python.
    pause
    exit /b 1
)

echo.
echo Starting service...
echo Submit page: http://127.0.0.1:8701/submit
echo Admin page: http://127.0.0.1:8701/admin
echo.
python -m uvicorn main:app --host 127.0.0.1 --port 8701
pause
