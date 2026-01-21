# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules

datas = []
hiddenimports = ['win32com.client', 'pythoncom']
datas += collect_data_files('customtkinter')
hiddenimports += collect_submodules('win32com')


a = Analysis(
    ['src\\app.py'],
    pathex=['C:\\Users\\darenes\\Desktop\\SPA1\\anSWer-Logging-Hub\\src'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='anSWer Logging Hub',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['C:\\Users\\darenes\\Desktop\\SPA1\\anSWer-Logging-Hub\\src\\ico\\CANoe_Logging.ico'],
)
