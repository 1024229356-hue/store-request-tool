@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PORT=8701"
cd /d "%~dp0" || goto cd_failed

echo.
echo ==============================
echo Store Request Tool startup
echo ==============================
echo CURRENT_DIR=%CD%
echo PORT=%PORT%
echo GIT_COMMIT=reported by startup_check.py
echo MAIN_FILE=reported by startup_check.py
echo ROUTE_COUNT=reported by startup_check.py
echo MISSING_ROUTES=reported by startup_check.py

if exist ".venv\Scripts\python.exe" goto venv_ready

echo Creating local Python virtual environment...
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -m venv .venv
    goto check_venv
)

where python >nul 2>nul
if errorlevel 1 goto python_missing
python -m venv .venv

:check_venv
if errorlevel 1 goto venv_failed
if not exist ".venv\Scripts\python.exe" goto venv_failed

:venv_ready
set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
echo PYTHON_EXE=%PYTHON_EXE%

call ".venv\Scripts\activate.bat"
if errorlevel 1 goto activate_failed

echo Installing dependencies...
python -m pip install --upgrade pip
if errorlevel 1 goto pip_failed
python -m pip install -r requirements.txt
if errorlevel 1 goto requirements_failed

echo Running startup diagnostics...
python startup_check.py --format text
if errorlevel 1 goto startup_check_failed

set "PORT_PID="
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do set "PORT_PID=%%p"
if defined PORT_PID goto port_busy

echo.
echo Standard URLs:
echo Submit:        http://127.0.0.1:8701/submit
echo Query:         http://127.0.0.1:8701/query
echo Admin login:   http://127.0.0.1:8701/admin/login
echo Dashboard:     http://127.0.0.1:8701/admin/dashboard
echo Tickets:       http://127.0.0.1:8701/admin
echo Schedules:     http://127.0.0.1:8701/admin/schedules
echo Employees:     http://127.0.0.1:8701/admin/employees
echo Shift types:   http://127.0.0.1:8701/admin/shift-types
echo Embedded:      http://127.0.0.1:8701/admin/embedded-pages
echo Version:       http://127.0.0.1:8701/__version
echo Health:        http://127.0.0.1:8701/healthz
echo Route health:  http://127.0.0.1:8701/admin/route-health
echo ==============================
echo.
echo Starting service...
python -m uvicorn main:app --host 127.0.0.1 --port 8701
if errorlevel 1 goto uvicorn_failed
echo Service stopped.
pause
exit /b 0

:cd_failed
echo ERROR: Cannot enter project directory.
pause
exit /b 1

:python_missing
echo ERROR: Python was not found. Please install Python 3 and run this script again.
pause
exit /b 1

:venv_failed
echo ERROR: Failed to create .venv.
pause
exit /b 1

:activate_failed
echo ERROR: Failed to activate .venv.
pause
exit /b 1

:pip_failed
echo ERROR: Failed to upgrade pip.
pause
exit /b 1

:requirements_failed
echo ERROR: requirements.txt installation failed.
pause
exit /b 1

:startup_check_failed
echo WARNING: Critical routes missing. Do not continue startup.
echo ERROR: startup_check.py failed. Run run_debug.bat and inspect logs\startup.log for full details.
pause
exit /b 1

:port_busy
echo WARNING: Port %PORT% is already in use by PID %PORT_PID%.
echo PID=%PORT_PID%
tasklist /FI "PID eq %PORT_PID%"
echo If this is a stale python.exe from this project, you can close it or run:
echo taskkill /PID %PORT_PID% /F
pause
exit /b 1

:uvicorn_failed
echo ERROR: uvicorn failed to start.
echo Run run_debug.bat and inspect logs\startup.log for full traceback.
pause
exit /b 1
