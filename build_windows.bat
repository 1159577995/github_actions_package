@echo off
chcp 65001 >nul 2>&1
REM ============================================================
REM Excel处理器 - Windows 打包脚本
REM 用法: 直接双击本文件，或在命令行执行 build_windows.bat
REM 产物: dist\Excel处理器-Dify答案匹配工具-windows.exe
REM ============================================================
setlocal
cd /d "%~dp0"

echo ======================================
echo Excel处理器 - Windows 打包
echo ======================================
echo.

REM ---------- 1. 查找 Python ----------
set "PYTHON_CMD="
py -3 --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py -3"
) else (
    python --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo 错误: 未找到 Python 3，请先安装 Python 3（https://www.python.org/downloads/）
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

%PYTHON_CMD% --version

REM 检查 Python 版本（依赖 langgraph/langchain-core 需 3.10+）
%PYTHON_CMD% -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: Python 版本过低，需要 3.10 及以上版本。
    echo 请安装 Python 3.10+（https://www.python.org/downloads/）后重新运行。
    pause
    exit /b 1
)

REM ---------- 2. 创建虚拟环境并安装依赖 ----------
echo.
echo 步骤1: 创建虚拟环境 build_venv ...
if not exist build_venv (
    %PYTHON_CMD% -m venv build_venv
    if %errorlevel% neq 0 (
        echo 错误: 创建虚拟环境失败，请确认 Python 支持 venv
        pause
        exit /b 1
    )
)

if not exist build_venv\Scripts\activate.bat (
    echo 错误: 虚拟环境创建失败（缺少 build_venv\Scripts\activate.bat）
    echo 请删除 build_venv 目录后重试，或重新安装 Python
    pause
    exit /b 1
)

call build_venv\Scripts\activate.bat

echo 步骤2: 升级 pip 并安装依赖...
python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo 错误: pip 升级失败
    pause
    exit /b 1
)

pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 错误: 依赖包安装失败
    pause
    exit /b 1
)

pip install pyinstaller
if %errorlevel% neq 0 (
    echo 错误: PyInstaller 安装失败
    pause
    exit /b 1
)

REM ---------- 3. 执行打包 ----------
REM 校验 spec 文件存在（避免拷贝文件不全时出现 "Spec file not found"）
if not exist excel_processor.spec (
    echo 错误: 缺少 excel_processor.spec 文件。
    echo 请确认已从开发机拷贝该文件（详见 打包文档.md 的文件清单，共 8 个文件）。
    pause
    exit /b 1
)

echo.
echo 步骤3: 开始打包 ...
set "BUNDLE_NAME=Excel处理器-Dify答案匹配工具-windows"
pyinstaller excel_processor.spec --noconfirm
if %errorlevel% neq 0 (
    echo 错误: 打包失败
    pause
    exit /b 1
)

REM ---------- 4. 输出结果 ----------
echo.
echo ======================================
echo 打包完成！
echo ======================================
echo 产物位置: dist\Excel处理器-Dify答案匹配工具-windows.exe
echo.
echo 使用方法：双击 dist 目录下的 Excel处理器-Dify答案匹配工具-windows.exe 即可运行
echo.
echo 提示：
echo - 首次启动较慢属正常现象（PyInstaller 单文件自解压）
echo - 文件处理 Agent 的沙箱工作区将生成在产物同目录 agent_workspace\ 下
echo - 如被杀毒软件误报，请添加信任（PyInstaller 单文件打包常见误报）
echo ======================================

REM CI（GitHub Actions 等）环境下不等待按键，避免任务挂起；本地双击运行时暂停以便查看
if not defined CI (
    pause
)
