@echo off
chcp 65001 >nul
:: ================================================================
::  MedOps Launcher - Joanna's Digital Matrix
::  Project: Tender Compliance Check
:: ================================================================

title 阿九Joanna · 标书合规排雷引擎 (MedOps Engine)
color 0A

echo.
echo  =============================================================
echo   Welcome to Joanna's MedTech Solutions
echo   正在初始化引擎环境... (Initializing Environment...)
echo  =============================================================
echo.

:: 1. 切换到脚本所在目录 (防止路径错误)
cd /d "%~dp0"

:: 2. 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 未检测到 Python 环境！请先安装 Python。
    pause
    exit
)

:: 3. 启动 Streamlit
echo   正在启动图形化界面 (GUI)...
echo   请稍候，浏览器将自动打开...
echo.

streamlit run step02_app_gui.py

:: 4. 如果崩溃，暂停显示错误信息
if %errorlevel% neq 0 (
    echo.
    echo [CRITICAL ERROR] 程序意外退出，请检查上方报错信息。
    pause
)