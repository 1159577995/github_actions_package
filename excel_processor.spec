# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Excel Processor Tool

产物名通过环境变量 BUNDLE_NAME 控制（各平台打包脚本按系统/架构设置，默认不设即用通用名）。
可选注入图标：环境变量 ICON_PATH 指向图标文件（Windows 用 .ico，macOS 用 .icns）。
"""
import os

block_cipher = None

name = os.environ.get('BUNDLE_NAME', 'Excel处理器-Dify答案匹配工具')
icon_path = os.environ.get('ICON_PATH', '')
use_icon = bool(icon_path) and os.path.exists(icon_path)

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[(icon_path, '.')] if use_icon else [],
    # langgraph/langchain 在 excel_processor.py 的函数内延迟导入（build_agent_graph），
    # PyInstaller 静态分析无法发现，必须显式声明，否则打包后文件处理 Agent 功能崩溃
    hiddenimports=[
        'openpyxl', 'requests',
        'langgraph', 'langgraph.graph', 'langgraph.checkpoint.memory',
        'langchain_core',
    ],
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
