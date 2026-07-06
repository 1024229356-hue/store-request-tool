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

set "LOCAL_COMMIT=unknown"
rem Uses: git rev-parse --short HEAD
for /f "usebackq tokens=* delims=" %%g in (`"%GIT_EXE%" rev-parse --short HEAD 2^>nul`) do set "LOCAL_COMMIT=%%g"

set "RUNNING_COMMIT="
for /f "usebackq tokens=* delims=" %%v in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $response = Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:8701/__version' -TimeoutSec 5; $payload = ConvertFrom-Json -InputObject $response.Content; if ($payload.git_commit) { [string]$payload.git_commit } else { '__NO_COMMIT__' } } catch { '__SERVICE_NOT_STARTED__' }"`) do set "RUNNING_COMMIT=%%v"

echo.
echo ==============================
echo Store Request Tool runtime check
echo ==============================
echo LOCAL_COMMIT=%LOCAL_COMMIT%

if not defined RUNNING_COMMIT goto service_not_started
if "%RUNNING_COMMIT%"=="__SERVICE_NOT_STARTED__" goto service_not_started

echo RUNNING_COMMIT=%RUNNING_COMMIT%

if /I "%LOCAL_COMMIT%"=="%RUNNING_COMMIT%" (
    echo 运行态一致。
) else (
    echo 运行态不一致，请执行 restart.bat。
)
goto finish_success

:service_not_started
echo 服务未启动。
goto finish_success

:cd_failed
echo ERROR: Cannot enter project directory.
goto finish_failed

:finish_success
pause
exit /b 0

:finish_failed
pause
exit /b 1
