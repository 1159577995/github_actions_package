# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Excel Processor Tool

产物名通过环境变量 BUNDLE_NAME 控制（各平台打包脚本按系统/架构设置，默认不设即用通用名）。
可选注入图标：环境变量 ICON_PATH 指向图标文件（Windows 用 .ico，macOS 用 .icns）。
"""
import os
from PyInstaller.utils.hooks import collect_submodules, copy_metadata

block_cipher = None

name = os.environ.get('BUNDLE_NAME', 'Excel处理器-Dify答案匹配工具')
icon_path = os.environ.get('ICON_PATH', '')
use_icon = bool(icon_path) and os.path.exists(icon_path)
langgraph_hiddenimports = collect_submodules('langgraph')
langchain_core_hiddenimports = collect_submodules('langchain_core')
extra_datas = []
for dist_name in ('langgraph', 'langchain-core'):
    try:
        extra_datas += copy_metadata(dist_name)
    except Exception:
        pass

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=([(icon_path, '.')] if use_icon else []) + extra_datas,
    # Windows 单文件打包对延迟导入更敏感；这里显式收集 LangGraph/LangChain Core
    # 的子模块与包元数据，避免双击启动时在 Agent 图构建阶段导入失败。
    hiddenimports=[
        'openpyxl', 'requests',
    ] + langgraph_hiddenimports + langchain_core_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI应用，不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if use_icon else None,
)
