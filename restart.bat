@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PORT=8701"
cd /d "%~dp0" || goto cd_failed

set "GIT_EXE=git"
where git >nul 2>nul
if errorlevel 1 (
    if exist "C:\Program Files\Git\cmd\git.exe" set "GIT_EXE=C:\Program Files\Git\cmd\git.exe"
)

set "GIT_COMMIT=unknown"
rem Uses: git rev-parse --short HEAD
for /f "usebackq tokens=* delims=" %%g in (`"%GIT_EXE%" rev-parse --short HEAD 2^>nul`) do set "GIT_COMMIT=%%g"

echo.
echo ==============================
echo Store Request Tool restart
echo ==============================
echo CURRENT_DIR=%CD%
echo GIT_COMMIT=%GIT_COMMIT%
echo PORT=%PORT%
echo.
echo Stopping old service when it belongs to this project...
call "%~dp0stop.bat" --no-pause

echo.
echo Waiting 1 second before startup...
timeout /t 1 /nobreak >nul

echo.
echo Starting current code with run.bat...
echo Health:        http://127.0.0.1:8701/healthz
echo Version:       http://127.0.0.1:8701/__version
echo Admin login:   http://127.0.0.1:8701/admin/login
echo Dashboard:     http://127.0.0.1:8701/admin/dashboard
echo ==============================
echo.
call "%~dp0run.bat"
exit /b %ERRORLEVEL%

:cd_failed
echo ERROR: Cannot enter project directory.
pause
exit /b 1
