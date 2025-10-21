@echo off
chcp 65001 >nul
title mykLabs Streamlit Application Launcher

echo ===============================
echo    mykLabs Streamlit 应用启动器
echo ===============================

:: 检查虚拟环境是否存在
if not exist ".venv\Scripts\activate.bat" (
    echo 错误：未找到虚拟环境 .venv
    echo 请先创建虚拟环境：python -m venv .venv
    pause
    exit /b 1
)

:: 激活虚拟环境
echo 正在激活虚拟环境...
call .venv\Scripts\activate.bat

:: 检查是否在虚拟环境中
python -c "import sys; print('虚拟环境激活成功' if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix) else '虚拟环境激活失败')" >nul 2>&1
if %errorlevel% neq 0 (
    echo 警告：虚拟环境可能未正确激活
)

:: 检查是否安装了 streamlit
python -c "import streamlit" >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误：虚拟环境中未安装 streamlit
    echo 请安装：pip install streamlit
    pause
    exit /b 1
)

:: 检查webui.py文件是否存在
if not exist "./webui/webui.py" (
    echo 错误：未找到 webui.py 文件
    echo 请确保脚本在正确目录下运行
    pause
    exit /b 1
)

echo 虚拟环境激活成功！
echo 正在启动 Streamlit 应用...
echo 应用启动后将在浏览器中自动打开
echo 按 Ctrl+C 可停止服务

:: streamlit run webui.py 作为包运行
python3 -m streamlit run webui/webui.py

if %errorlevel% neq 0 (
    echo.
    echo 应用运行出现错误，返回代码：%errorlevel%
)

pause