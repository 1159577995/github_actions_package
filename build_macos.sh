#!/bin/bash
# ============================================================
# Excel处理器 - macOS 打包脚本
# 支持架构: arm64(Apple Silicon) / x86_64(Intel)
# 要求: Python 3.10+（依赖 langgraph/langchain-core 需要）
# 用法: ./build_macos.sh   （或由 build_all.sh 自动调用）
# ============================================================
set -e

cd "$(dirname "$0")"

echo "======================================"
echo "Excel处理器 - macOS 打包"
echo "======================================"
echo ""

# ---------- 1. 系统检测 ----------
ARCH="$(uname -m)"
case "$ARCH" in
    arm64) ARCH_LABEL="arm64" ;;
    x86_64) ARCH_LABEL="x86_64" ;;
    *) ARCH_LABEL="$ARCH" ;;
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
    echo "  请从 https://www.python.org/downloads/ 下载安装 Python 3.10+，"
    echo "  安装后重新运行本脚本。"
    exit 1
fi

echo "使用 Python: $PY_BIN ($($PY_BIN --version))"

# ---------- 3. 创建虚拟环境并安装依赖 ----------
echo ""
echo "步骤1: 创建虚拟环境 build_venv ..."

if [ ! -d "build_venv" ]; then
    "$PY_BIN" -m venv build_venv
fi

if [ ! -f "build_venv/bin/activate" ]; then
    echo "错误: 虚拟环境创建失败（build_venv/bin/activate 不存在）。"
    echo "  请确认 Python 安装完整（包含 venv/ensurepip），或删除 build_venv 目录后重试。"
    exit 1
fi
source build_venv/bin/activate

echo "步骤2: 升级 pip 并安装依赖..."
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

# ---------- 4. 执行打包 ----------
# 校验 spec 文件存在（避免拷贝文件不全时出现 "Spec file not found"）
if [ ! -f "excel_processor_macos.spec" ]; then
    echo "错误: 缺少 excel_processor_macos.spec 文件。"
    echo "  请确认已从开发机拷贝该文件（详见 打包文档.md 的文件清单，共 8 个文件）。"
    exit 1
fi

echo ""
echo "步骤3: 开始打包 (架构 $ARCH_LABEL) ..."
# 使用 excel_processor_macos.spec 生成 .app 应用包：双击直接出界面、无命令行、免自解压快速启动
BUNDLE_NAME="Excel处理器-Dify答案匹配工具" pyinstaller excel_processor_macos.spec --noconfirm

BINARY="dist/Excel处理器-Dify答案匹配工具.app"

# ---------- 5. 输出结果 ----------
echo ""
echo "======================================"
echo "打包完成！"
echo "======================================"
echo "产物位置: $BINARY"
echo ""
echo "使用方法（双击运行）："
echo "1. 在访达中进入 dist 目录"
echo "2. 双击 Excel处理器-Dify答案匹配工具.app 即可运行（无命令行窗口，含启动画面）"
echo ""
echo "提示:"
echo "- 本地构建的应用可直接双击运行，无 Gatekeeper 拦截"
echo "- 若拷贝到其他 Mac 提示“无法打开，因为来自身份不明的开发者”，"
echo "  请右键 → 打开，或执行: xattr -dr com.apple.quarantine \"$BINARY\""
echo "- 文件处理 Agent 的沙箱工作区将生成在 dist/agent_workspace/ 下"
echo "======================================"
