@echo off
echo Starting Smart Transfer Optimization System...
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found. Please install Python 3.8+ first.
    pause
    exit /b 1
)

REM 安装依赖包
echo Installing dependencies...
pip install -r requirements.txt

if errorlevel 1 (
    echo Error: Failed to install dependencies.
    pause
    exit /b 1
)

REM 启动Web应用
echo Starting web interface...
streamlit run web_interface.py

pause