@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PORT=8701"
set "NO_PAUSE="
if /I "%~1"=="--no-pause" set "NO_PAUSE=1"
cd /d "%~dp0" || goto cd_failed

echo.
echo ==============================
echo Store Request Tool stop
echo ==============================
echo CURRENT_DIR=%CD%
echo PORT=%PORT%

set "PORT_PID="
for /f "usebackq tokens=* delims=" %%p in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$connections = @(Get-NetTCPConnection -LocalPort %PORT% -State Listen -ErrorAction SilentlyContinue); if ($connections.Count -gt 0) { [string]$connections[0].OwningProcess }"`) do set "PORT_PID=%%p"

if not defined PORT_PID goto not_running

echo 当前 8701 已被 PID %PORT_PID% 占用。
powershell -NoProfile -ExecutionPolicy Bypass -Command "$process = Get-CimInstance Win32_Process -Filter 'ProcessId=%PORT_PID%' -ErrorAction SilentlyContinue; if ($process -and $process.CommandLine) { Write-Host ('CommandLine: ' + $process.CommandLine) } else { Write-Host 'CommandLine: [无法读取]' }"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$process = Get-CimInstance Win32_Process -Filter 'ProcessId=%PORT_PID%' -ErrorAction SilentlyContinue; $cmd = if ($process) { [string]$process.CommandLine } else { '' }; if ($cmd -like '*store_request_tool*' -or $cmd -like '*uvicorn main:app*' -or $cmd -like '*D:\需求小程序\store_request_tool*') { exit 0 } else { exit 1 }"
if errorlevel 1 goto not_project_process

echo 确认是本项目进程，正在停止 PID %PORT_PID%...
taskkill /PID %PORT_PID% /F
if errorlevel 1 goto kill_failed
echo 已停止 PID %PORT_PID%。
goto finish_success

:not_running
echo 8701 未运行。
goto finish_success

:not_project_process
echo 无法确认这是本项目进程，不会自动停止。
echo 请根据上面的 PID 和 CommandLine 手动确认后再处理。
goto finish_success

:kill_failed
echo ERROR: 停止 PID %PORT_PID% 失败，请手动检查。
goto finish_failed

:cd_failed
echo ERROR: Cannot enter project directory.
goto finish_failed

:finish_success
if defined NO_PAUSE exit /b 0
pause
exit /b 0

:finish_failed
if defined NO_PAUSE exit /b 1
pause
exit /b 1
