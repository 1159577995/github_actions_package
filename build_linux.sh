#!/bin/bash
# ============================================================
# Excel处理器 - Linux / 信创(麒麟Kylin、统信UOS) 打包脚本
# 支持架构: x86_64 / aarch64(鲲鹏、飞腾)
# 要求: Python 3.10+（依赖 langgraph/langchain-core 需要）
# 用法: ./build_linux.sh   （或由 build_all.sh 自动调用）
# ============================================================
set -e

cd "$(dirname "$0")"

echo "======================================"
echo "Excel处理器 - Linux/信创 打包"
echo "======================================"
echo ""

# ---------- 1. 系统检测 ----------
ARCH="$(uname -m)"
case "$ARCH" in
    x86_64)   ARCH_LABEL="x86_64" ;;
    aarch64|arm64) ARCH_LABEL="aarch64" ;;
    *)
        echo "警告: 未识别的架构 $ARCH，产物将使用 $ARCH 作为后缀"
        ARCH_LABEL="$ARCH"
        ;;
esac
echo "架构: $ARCH_LABEL"

# ---------- 2. 寻找满足版本要求的 Python (>= 3.10) ----------
PY_BIN=""
for cand in python3 python3.13 python3.12 python3.11 python3.10; do
    if command -v "$cand" &> /dev/null; then
        MAJOR=$("$cand" -c "import sys; print(sys.version_info.major)")
        MINOR=$("$cand" -c "import sys; print(sys.version_info.minor)")
        if [ "$MAJOR" -gt 3 ] || { [ "$MAJOR" -eq 3 ] && [ "$MINOR" -ge 10 ]; }; then
            PY_BIN="$cand"
            break
        fi
    fi
done

if [ -z "$PY_BIN" ]; then
    echo "错误: 未找到 Python 3.10 及以上版本。"
    echo "  本程序依赖（langgraph / langchain-core）要求 Python ≥ 3.10，"
    echo "  系统自带 Python 3.8 无法打包。"
    echo "  请安装更高版本 Python，例如："
    echo "    sudo apt update"
    echo "    sudo apt install -y python3.11 python3.11-venv python3.11-tk"
    echo "  （麒麟/统信等信创系统请以软件源实际提供的版本为准，如 python3.10/python3.11）"
    echo "  安装完成后重新运行本脚本，脚本会自动选用高版本 Python。"
    exit 1
fi

echo "使用 Python: $PY_BIN ($($PY_BIN --version))"

# 检查 tkinter（打包 GUI 必需）
if ! "$PY_BIN" -c "import tkinter" &> /dev/null; then
    echo "错误: $PY_BIN 缺少 tkinter 支持，无法打包 GUI 程序。"
    echo "  例如: sudo apt install python3.11-tk"
    echo "  （包名中的版本号需与所用 Python 一致）"
    exit 1
fi

# ---------- 3. 创建虚拟环境并安装依赖 ----------
echo ""
echo "步骤1: 创建虚拟环境 build_venv ..."

if [ ! -d "build_venv" ]; then
    "$PY_BIN" -m venv build_venv
fi

# 校验 venv 是否真正创建成功（部分系统缺 ensurepip 会创建不完整目录）
if [ ! -f "build_venv/bin/activate" ]; then
    echo "警告: python3 -m venv 创建失败（常见原因：缺少 python3-venv / ensurepip）。"
    echo "  尝试使用 virtualenv 回退..."
    if command -v pip3 &> /dev/null; then
        pip3 install --user virtualenv >/dev/null 2>&1 || true
    fi
    if "$PY_BIN" -m virtualenv --version &> /dev/null; then
        rm -rf build_venv
        "$PY_BIN" -m virtualenv build_venv
    else
        echo "错误: venv 与 virtualenv 均不可用，请先安装 python3-venv："
        echo "  sudo apt install -y python3.11-venv"
        echo "  （或: sudo apt install -y python3-virtualenv）"
        exit 1
    fi
fi

if [ ! -f "build_venv/bin/activate" ]; then
    echo "错误: 虚拟环境创建仍然失败，无法继续打包。"
    exit 1
fi
source build_venv/bin/activate

echo "步骤2: 升级 pip 并安装依赖..."
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

# ---------- 4. 执行打包 ----------
# 校验 spec 文件存在（避免拷贝文件不全时出现 "Spec file not found"）
if [ ! -f "excel_processor.spec" ]; then
    echo "错误: 缺少 excel_processor.spec 文件。"
    echo "  请确认已从开发机拷贝该文件（详见 打包文档.md 的文件清单，共 8 个文件）。"
    exit 1
fi

echo ""
echo "步骤3: 开始打包 (架构 $ARCH_LABEL) ..."
BUNDLE_NAME="Excel处理器-Dify答案匹配工具-linux-${ARCH_LABEL}" pyinstaller excel_processor.spec --noconfirm

# ---------- 5. 输出结果 ----------
BINARY="dist/Excel处理器-Dify答案匹配工具-linux-${ARCH_LABEL}"
if [ -f "$BINARY" ]; then
    chmod +x "$BINARY"
fi

echo ""
echo "======================================"
echo "打包完成！"
echo "======================================"
echo "产物位置: $BINARY"
echo ""
echo "使用方法（双击运行）："
echo "1. 在文件管理器中进入 dist 目录"
echo "2. 双击 $BINARY 即可运行"
echo "   若提示无执行权限，请在终端执行: chmod +x \"$BINARY\""
echo ""
echo "提示:"
echo "- 首次启动较慢属正常现象（PyInstaller 单文件自解压）"
echo "- 文件处理 Agent 的沙箱工作区优先生成在产物同目录 agent_workspace/ 下；若同级目录不可写，则回退到 ~/.Excel处理器-Dify答案匹配工具/agent_workspace/"
echo "======================================"
