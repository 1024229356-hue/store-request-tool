@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo 正在创建本地 Python 虚拟环境...
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if errorlevel 1 (
        echo 创建虚拟环境失败，请先安装 Python 3。
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖安装失败，请检查网络或 Python 环境。
    pause
    exit /b 1
)

echo.
echo 服务启动中...
echo 门店提报页: http://127.0.0.1:8701/submit
echo 后台管理页: http://127.0.0.1:8701/admin
echo.
python -m uvicorn main:app --host 127.0.0.1 --port 8701
pause
