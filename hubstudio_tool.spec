# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置。"""

from PyInstaller.utils.hooks import collect_all

block_cipher = None

playwright_datas, playwright_binaries, playwright_hidden = collect_all("playwright")

a = Analysis(
    ["run_app.py"],
    pathex=[],
    binaries=playwright_binaries,
    datas=[
        ("templates", "templates"),
        ("config.example.json", "."),
        ("使用说明.txt", "."),
    ] + playwright_datas,
    hiddenimports=[
        "flask",
        "werkzeug",
        "jinja2",
        "openpyxl",
        "win32api",
        "win32con",
        "win32gui",
        "win32process",
        "win32com",
        "win32com.client",
        "pythoncom",
        "pywintypes",
        "excel_sync",
        "hubstudio_client",
        "app_paths",
        "web_app",
        "outlook_checker",
        "foxmail_automation",
        "recovery_email",
    ] + playwright_hidden,
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
    [],
    exclude_binaries=True,
    name="HubStudio批量控制台",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="HubStudio批量控制台",
)
