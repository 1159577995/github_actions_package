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
    """最早阶段的日志：写入可执行文件同目录的 运行日志.log"""
    try:
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(os.path.abspath(sys.executable))
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(base, '运行日志.log'), 'a', encoding='utf-8') as fh:
            fh.write(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
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
