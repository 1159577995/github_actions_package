#!/bin/bash
# ============================================================
# Excel处理器 - 跨平台一键打包主入口
# 支持: macOS / Linux / 信创(麒麟Kylin、统信UOS，x86_64 与 aarch64)
# Windows 请直接双击运行 build_windows.bat
#
# 用法:
#   ./build_all.sh
# 注意:
#   PyInstaller 不支持交叉编译，请在目标系统上执行本脚本，
#   信创电脑请在信创机器上运行，脚本会自动识别架构。
# ============================================================
set -e

# 切换到脚本所在目录（项目根）
cd "$(dirname "$0")"

echo "======================================"
echo "Excel处理器 - 跨平台一键打包"
echo "当前系统: $(uname -s) / $(uname -m)"
echo "======================================"
echo ""

case "$(uname -s)" in
    Darwin)
        echo "检测到 macOS，开始打包..."
        bash build_macos.sh
        ;;
    Linux)
        echo "检测到 Linux（含信创麒麟/统信），开始打包..."
        bash build_linux.sh
        ;;
    *)
        echo "错误: 不支持的系统 $(uname -s)"
        echo "Windows 请直接双击 build_windows.bat"
        exit 1
        ;;
esac
