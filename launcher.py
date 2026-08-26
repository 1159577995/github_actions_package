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

APP_NAME = "Excel处理器-Dify答案匹配工具"


def _is_app_translocated(path):
    """macOS 通过 App Translocation 运行时，应用会落到只读的临时挂载目录。"""
    norm = os.path.abspath(path)
    return sys.platform == 'darwin' and '/AppTranslocation/' in norm


def _is_writable_dir(path):
    """检测目录是否可写；不存在时尝试创建并写一个临时探针文件。"""
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, '.write_probe')
        with open(probe, 'a', encoding='utf-8'):
            pass
        os.remove(probe)
        return True
    except Exception:
        return False


def _default_user_data_dir():
    """跨平台用户数据目录：当产物同级目录不可写时，回退到这里。"""
    home = os.path.expanduser('~')
    if sys.platform == 'darwin':
        return os.path.join(home, 'Library', 'Application Support', APP_NAME)
    if os.name == 'nt':
        return os.path.join(os.environ.get('APPDATA', home), APP_NAME)
    return os.path.join(home, f'.{APP_NAME}')


def _resolve_base_dir():
    """解析产物同级目录：macOS .app 提升到 .app 同级，其他平台取可执行文件同级。"""
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        if os.path.basename(exe_dir) == 'MacOS' and os.path.basename(os.path.dirname(exe_dir)) == 'Contents':
            return os.path.dirname(os.path.dirname(os.path.dirname(exe_dir)))
        return exe_dir
    return os.path.dirname(os.path.abspath(__file__))


_LOG_DIR_CACHE = [None]


def _launcher_log(msg):
    """最早阶段的日志：优先写入产物同级目录，不可写时回退到跨平台用户数据目录。"""
    try:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n"
        if _LOG_DIR_CACHE[0] is None:
            base = _resolve_base_dir()
            candidates = []
            if not _is_app_translocated(base):
                candidates.append(base)
            candidates.append(_default_user_data_dir())
            candidates.append(os.path.expanduser('~'))
            for d in candidates:
                if _is_writable_dir(d):
                    _LOG_DIR_CACHE[0] = d
                    break
            if _LOG_DIR_CACHE[0] is None:
                _LOG_DIR_CACHE[0] = os.path.expanduser('~')
        with open(os.path.join(_LOG_DIR_CACHE[0], '运行日志.log'), 'a', encoding='utf-8') as fh:
            fh.write(line)
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
