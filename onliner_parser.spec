# -*- mode: python ; coding: utf-8 -*-
import os

cwd_path = os.getcwd()


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

icon_res_path = 'app\\source\\onliner_parser.ico'
icon_abs_path = os.path.join(cwd_path, icon_res_path)
a.datas.append((icon_res_path, icon_abs_path, 'DATA'))
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Onliner Parser',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_res_path,
)
