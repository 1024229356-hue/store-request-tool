@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PORT=8701"
cd /d "%~dp0" || goto cd_failed

if not exist "logs" mkdir "logs"
set "STARTUP_LOG=logs\startup.log"

echo Startup debug log: %CD%\%STARTUP_LOG%
echo ===== startup debug %DATE% %TIME% ===== > "%STARTUP_LOG%"
echo CURRENT_DIR=%CD% >> "%STARTUP_LOG%" 2>&1
echo PORT=%PORT% >> "%STARTUP_LOG%" 2>&1

if exist ".venv\Scripts\python.exe" goto venv_ready

echo Creating local Python virtual environment... >> "%STARTUP_LOG%" 2>&1
where py >> "%STARTUP_LOG%" 2>&1
if not errorlevel 1 (
    py -3 -m venv .venv >> "%STARTUP_LOG%" 2>&1
    goto check_venv
)

where python >> "%STARTUP_LOG%" 2>&1
if errorlevel 1 goto python_missing
python -m venv .venv >> "%STARTUP_LOG%" 2>&1

:check_venv
if errorlevel 1 goto venv_failed
if not exist ".venv\Scripts\python.exe" goto venv_failed

:venv_ready
set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
echo PYTHON_EXE=%PYTHON_EXE% >> "%STARTUP_LOG%" 2>&1

call ".venv\Scripts\activate.bat" >> "%STARTUP_LOG%" 2>&1
if errorlevel 1 goto activate_failed

echo Installing dependencies... >> "%STARTUP_LOG%" 2>&1
python -m pip install --upgrade pip >> "%STARTUP_LOG%" 2>&1
if errorlevel 1 goto pip_failed
python -m pip install -r requirements.txt >> "%STARTUP_LOG%" 2>&1
if errorlevel 1 goto requirements_failed

echo Running startup diagnostics... >> "%STARTUP_LOG%" 2>&1
python startup_check.py --format text >> "%STARTUP_LOG%" 2>&1
if errorlevel 1 goto startup_check_failed

set "PORT_PID="
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do set "PORT_PID=%%p"
if defined PORT_PID goto port_busy
echo PORT_8701=free >> "%STARTUP_LOG%" 2>&1

echo Standard URLs: >> "%STARTUP_LOG%" 2>&1
echo http://127.0.0.1:8701/submit >> "%STARTUP_LOG%" 2>&1
echo http://127.0.0.1:8701/query >> "%STARTUP_LOG%" 2>&1
echo http://127.0.0.1:8701/admin/login >> "%STARTUP_LOG%" 2>&1
echo http://127.0.0.1:8701/admin/dashboard >> "%STARTUP_LOG%" 2>&1
echo http://127.0.0.1:8701/healthz >> "%STARTUP_LOG%" 2>&1
echo http://127.0.0.1:8701/__version >> "%STARTUP_LOG%" 2>&1
echo Starting uvicorn... >> "%STARTUP_LOG%" 2>&1
python -m uvicorn main:app --host 127.0.0.1 --port 8701 >> "%STARTUP_LOG%" 2>&1
if errorlevel 1 goto uvicorn_failed
echo Service stopped. >> "%STARTUP_LOG%" 2>&1
type "%STARTUP_LOG%"
pause
exit /b 0

:cd_failed
echo ERROR: Cannot enter project directory.
pause
exit /b 1

:python_missing
echo ERROR: Python was not found. >> "%STARTUP_LOG%" 2>&1
goto fail

:venv_failed
echo ERROR: Failed to create .venv. >> "%STARTUP_LOG%" 2>&1
goto fail

:activate_failed
echo ERROR: Failed to activate .venv. >> "%STARTUP_LOG%" 2>&1
goto fail

:pip_failed
echo ERROR: Failed to upgrade pip. >> "%STARTUP_LOG%" 2>&1
goto fail

:requirements_failed
echo ERROR: requirements.txt installation failed. >> "%STARTUP_LOG%" 2>&1
goto fail

:startup_check_failed
echo ERROR: startup_check.py failed. >> "%STARTUP_LOG%" 2>&1
goto fail

:port_busy
echo WARNING: Port %PORT% is already in use by PID %PORT_PID%. >> "%STARTUP_LOG%" 2>&1
echo PID=%PORT_PID% >> "%STARTUP_LOG%" 2>&1
tasklist /FI "PID eq %PORT_PID%" >> "%STARTUP_LOG%" 2>&1
echo If this is a stale python.exe from this project, you can close it or run: >> "%STARTUP_LOG%" 2>&1
echo taskkill /PID %PORT_PID% /F >> "%STARTUP_LOG%" 2>&1
goto fail

:uvicorn_failed
echo ERROR: uvicorn failed to start. >> "%STARTUP_LOG%" 2>&1
goto fail

:fail
type "%STARTUP_LOG%"
echo.
echo Startup failed. Full log: %CD%\%STARTUP_LOG%
pause
exit /b 1
