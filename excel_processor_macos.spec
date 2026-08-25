# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file - macOS .app 打包（onedir + BUNDLE）

产物: dist/<BUNDLE_NAME>.app
特点: 双击直接出界面（无命令行窗口）、免自解压快速启动、可注入启动画面。
产物名通过环境变量 BUNDLE_NAME 控制；可选注入图标 ICON_PATH（.icns）。
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
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,  # onedir：二进制分离，加快启动
    name=name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 应用，不显示控制台/终端
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if use_icon else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=name,
)

app = BUNDLE(
    coll,
    name=name + '.app',
    icon=icon_path if use_icon else None,
    bundle_identifier='com.excelprocessor.difytool',
    info_plist={
        'NSHighResolutionCapable': True,
        'CFBundleDisplayName': 'Excel处理器',
        'CFBundleName': name,
    },
)
