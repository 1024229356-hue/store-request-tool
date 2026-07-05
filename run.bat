@echo off
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
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
echo Runtime diagnostics:
echo CURRENT_DIR=%CD%
set "GIT_EXE=git"
where git >nul 2>nul
if errorlevel 1 (
    if exist "C:\Program Files\Git\cmd\git.exe" set "GIT_EXE=C:\Program Files\Git\cmd\git.exe"
)
set "GIT_COMMIT=unknown"
for /f "usebackq tokens=*" %%i in (`"%GIT_EXE%" rev-parse --short HEAD 2^>nul`) do set "GIT_COMMIT=%%i"
echo GIT_COMMIT=%GIT_COMMIT%
python -c "import sys; print('PYTHON_EXE=', sys.executable)"
python -c "import main; required=['/admin/my-work','/admin/archive','/admin/trash','/admin/employees','/admin/shift-types','/admin/schedules','/admin/tickets/bulk-archive','/admin/tickets/bulk-delete']; paths=[r.path for r in main.app.routes]; missing=[p for p in required if p not in paths]; print('MAIN_FILE=', main.__file__); print('ROUTE_COUNT=', len(main.app.routes)); print('MISSING_ROUTES=', missing); raise SystemExit(1 if missing else 0)"
if errorlevel 1 (
    echo Critical routes missing. Do not continue startup.
    pause
    exit /b 1
)

echo.
echo Starting service...
echo Submit page: http://127.0.0.1:8701/submit
echo Admin page: http://127.0.0.1:8701/admin
echo Version: http://127.0.0.1:8701/__version
echo.
python -m uvicorn main:app --host 127.0.0.1 --port 8701
pause
