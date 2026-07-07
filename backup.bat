@echo off
setlocal

cd /d "%~dp0"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "STAMP=%%i"
set "BACKUP_DIR=%CD%\backups\%STAMP%"

mkdir "%BACKUP_DIR%\data" >nul 2>nul

if exist "data\tickets.db" (
  copy /Y "data\tickets.db" "%BACKUP_DIR%\data\tickets.db" >nul
) else (
  echo [WARN] data\tickets.db not found.
)

if exist "data\embedded_pages" (
  xcopy "data\embedded_pages" "%BACKUP_DIR%\data\embedded_pages\" /E /I /Y >nul
) else (
  echo [WARN] data\embedded_pages not found.
)

if exist "uploads" (
  xcopy "uploads" "%BACKUP_DIR%\uploads\" /E /I /Y >nul
) else (
  echo [WARN] uploads not found.
)

echo Backup completed:
echo %BACKUP_DIR%

endlocal
