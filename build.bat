@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: ============================================================
::  巡检数据处理工具 - Windows 打包脚本
::  用法: 双击 build.bat 运行
:: ============================================================

echo ==================================================
echo   巡检数据处理工具 - 打包脚本
echo ==================================================
echo.

:: ============ 配置区 ============
set "SCRIPT_FILE=mergeExcel.py"
set "OUTPUT_NAME=巡检数据处理工具"
set "ICON_FILE=images\favicon.ico"
set "BUILD_DIR=venv"

:: ============ 检查 Python 环境 ============
echo [1/8] 检查 Python 环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python
    pause
    exit /b 1
)
python --version

:: ============ 检查入口文件 ============
echo.
echo [2/8] 检查入口文件...
if not exist "%SCRIPT_FILE%" (
    echo [错误] 找不到入口文件: %SCRIPT_FILE%
    pause
    exit /b 1
)
echo 入口文件: %SCRIPT_FILE%

:: ============ 检查图标文件 ============
echo.
echo [3/8] 检查图标文件...
if not exist "%ICON_FILE%" (
    echo [警告] 找不到图标文件: %ICON_FILE%，将使用默认图标
    set "ICON_PARAM="
) else (
    echo 图标文件: %ICON_FILE%
    set "ICON_PARAM=-i %ICON_FILE%"
)

:: ============ 创建虚拟环境 ============
echo.
echo [4/8] 准备虚拟环境...
if not exist "%BUILD_DIR%" (
    echo 创建虚拟环境: %BUILD_DIR%
    python -m venv %BUILD_DIR%
) else (
    echo 虚拟环境已存在，跳过创建
)

:: ============ 激活虚拟环境 ============
echo.
echo [5/8] 激活虚拟环境...
call %BUILD_DIR%\Scripts\activate.bat

:: ============ 安装依赖 ============
echo.
echo [6/8] 安装必要依赖...
pip install --upgrade pip -q
pip install -q openpyxl paramiko pythonping PyYAML alive-progress pyinstaller

:: ============ 清理旧的构建文件 ============
echo.
echo [7/8] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
if exist *.spec del /q *.spec

:: ============ 打包 ============
echo.
echo [8/8] 开始打包...
echo.

pyinstaller ^
    -F ^
    %ICON_PARAM% ^
    -n "%OUTPUT_NAME%" ^
    --collect-all grapheme ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module PyQt5 ^
    --exclude-module PyQt6 ^
    --exclude-module tkinter ^
    --exclude-module test ^
    --exclude-module unittest ^
    --exclude-module email ^
    --exclude-module xml ^
    --exclude-module pydoc ^
    --exclude-module doctest ^
    --clean ^
    %SCRIPT_FILE%

:: ============ 检查结果 ============
echo.
if exist "dist\%OUTPUT_NAME%.exe" (
    echo ==================================================
    echo   打包成功！
    echo ==================================================
    echo.
    echo   输出文件: dist\%OUTPUT_NAME%.exe
    echo.
    dir "dist\%OUTPUT_NAME%.exe"
    echo.
    echo ==================================================
) else (
    echo ==================================================
    echo   打包失败，请检查错误信息
    echo ==================================================
)

pause
