# -*- mode: python ; coding: utf-8 -*-

import openpyxl
import os

block_cipher = None

openpyxl_path = os.path.dirname(openpyxl.__file__)

a = Analysis(
    ['sub2.py'],
    pathex=[],
    binaries=[],
    datas=[(openpyxl_path, 'openpyxl')],
    hiddenimports=['openpyxl'],
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
    name='sub2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
