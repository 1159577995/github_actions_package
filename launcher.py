#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
程序启动入口（PyInstaller 打包用，excel_processor.spec 与 excel_processor_macos.spec 的入口脚本）。

作用：在加载主程序之前就开始写日志——只要双击运行就会在可执行文件同目录生成/追加
`运行日志.log`，即使主程序加载或启动阶段失败，也能记录异常信息，避免“双击无反应”。
"""
import os
import sys
import traceback
from datetime import datetime


def _launcher_log(msg):
    """最早阶段的日志：写入可执行文件（或 .app 应用包）同级目录的 运行日志.log，
    同级目录不可写（如 Program Files）时回退到用户主目录。"""
    try:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n"
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(os.path.abspath(sys.executable))
            # macOS .app 包内可执行文件路径为 App.app/Contents/MacOS/App，
            # 提升到 .app 同级目录（dist/），与主程序 BASE_DIR 保持一致
            if os.path.basename(exe_dir) == 'MacOS' and os.path.basename(os.path.dirname(exe_dir)) == 'Contents':
                base = os.path.dirname(os.path.dirname(os.path.dirname(exe_dir)))
            else:
                base = exe_dir
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        for d in (base, os.path.expanduser('~')):
            try:
                with open(os.path.join(d, '运行日志.log'), 'a', encoding='utf-8') as fh:
                    fh.write(line)
                break  # 写入成功即停止
            except Exception:
                continue
    except Exception:
        pass


def main():
    _launcher_log("=" * 60)
    _launcher_log("程序启动")
    try:
        import excel_processor
        excel_processor.main()  # 主程序内部同样全程写日志
    except Exception:
        _launcher_log("启动异常:\n" + traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
