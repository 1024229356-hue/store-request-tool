@echo off
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PORT=8701"
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

set "GIT_EXE="
where git >nul 2>nul
if %errorlevel%==0 set "GIT_EXE=git"
if not defined GIT_EXE if exist "C:\Program Files\Git\cmd\git.exe" set "GIT_EXE=C:\Program Files\Git\cmd\git.exe"
if not defined GIT_EXE if exist "C:\Program Files\Git\bin\git.exe" set "GIT_EXE=C:\Program Files\Git\bin\git.exe"
set "GIT_COMMIT=unknown"
if defined GIT_EXE (
    for /f "usebackq tokens=*" %%i in (`"%GIT_EXE%" rev-parse --short HEAD 2^>nul`) do set "GIT_COMMIT=%%i"
)

set "PORT_PID="
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
    if not defined PORT_PID set "PORT_PID=%%p"
)

echo.
echo ==============================
echo 止痒 ERP 本地服务启动信息
echo ==============================
echo 项目目录：%CD%
echo CURRENT_DIR=%CD%
echo 当前版本：%GIT_COMMIT%
echo 端口：%PORT%
echo PORT=%PORT%
python -c "import sys; print('Python 路径：' + sys.executable); print('PYTHON_EXE=' + sys.executable)"
python -c "import main, sys; missing=main.required_missing_routes(main.app); print('main.py 路径：' + str(main.BASE_DIR / 'main.py')); print('MAIN_FILE=' + str(main.BASE_DIR / 'main.py')); print('ROUTE_COUNT=' + str(len(main.app.routes))); print('关键路由缺失情况：' + ('无' if not missing else ', '.join(missing))); print('MISSING_ROUTES=' + ('[]' if not missing else ','.join(missing))); sys.exit(1 if missing else 0)"
if errorlevel 1 (
    echo WARNING: Critical routes missing. Do not continue startup.
    pause
    exit /b 1
)

if defined PORT_PID (
    echo WARNING: Port %PORT% is already in use by PID %PORT_PID%.
    echo PID=%PORT_PID%
    tasklist /FI "PID eq %PORT_PID%"
    pause
    exit /b 1
)

echo.
echo 门店提交：
echo http://127.0.0.1:8701/submit
echo.
echo 门店查询：
echo http://127.0.0.1:8701/query
echo.
echo 后台登录：
echo http://127.0.0.1:8701/admin/login
echo.
echo 业务总览：
echo http://127.0.0.1:8701/admin/dashboard
echo.
echo 工单管理：
echo http://127.0.0.1:8701/admin
echo.
echo 门店排班：
echo http://127.0.0.1:8701/admin/schedules
echo.
echo 员工管理：
echo http://127.0.0.1:8701/admin/employees
echo.
echo 班次设置：
echo http://127.0.0.1:8701/admin/shift-types
echo.
echo 嵌入页面管理：
echo http://127.0.0.1:8701/admin/embedded-pages
echo.
echo 运行版本：
echo http://127.0.0.1:8701/__version
echo.
echo 健康检查：
echo http://127.0.0.1:8701/healthz
echo.
echo 路由体检：
echo http://127.0.0.1:8701/admin/route-health
echo ==============================
echo.

echo Starting service...
python -m uvicorn main:app --host 127.0.0.1 --port 8701
pause
