@echo off
chcp 65001 >nul
title BOM生成工具 v1.0

echo.
echo =====================================
echo      BOM生成工具 v1.0 启动器
echo =====================================
echo.

echo 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python环境
    echo 请确保已安装Python 3.10+并添加到系统PATH
    pause
    exit /b 1
)

echo Python环境检查通过
echo.

echo 启动BOM生成工具...
echo.

REM 优先使用单文件版本
if exist "BOM_Generator_v1.0.py" (
    echo 使用单文件版本启动...
    python BOM_Generator_v1.0.py
) else if exist "src\main.py" (
    echo 使用模块化版本启动...
    python src\main.py
) else (
    echo 错误: 未找到应用程序文件
    echo 请确保 BOM_Generator_v1.0.py 或 src\main.py 文件存在
    pause
    exit /b 1
)

echo.
echo 应用程序已退出
pause
