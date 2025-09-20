@echo off
echo Markdown流水转Excel工具启动器
echo ==============================

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.7或更高版本
    pause
    exit /b 1
)

REM 检查依赖包
echo 正在检查依赖包...
python -c "import pandas, openpyxl, yaml" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖包...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo 依赖包安装失败，请手动运行: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

echo 启动工具...
python markdown_to_excel.py

pause