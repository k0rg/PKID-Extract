# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PKID Extract Tool.

Builds a single-file Windows executable that bundles:
  - extract_PKID.py  (the GUI application)
  - oa3tool.exe      (the HW-hash decoder)

Usage (from the repo root):
    pyinstaller PKID_Extract.spec
"""

import os

block_cipher = None
script_dir = os.path.dirname(os.path.abspath(SPECPATH))

a = Analysis(
    [os.path.join(script_dir, 'extract_PKID.py')],
    pathex=[script_dir],
    binaries=[],
    datas=[
        # Bundle oa3tool.exe so it is extracted at runtime into sys._MEIPASS
        (os.path.join(script_dir, 'oa3tool.exe'), '.'),
    ],
    hiddenimports=['openpyxl'],
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
    name='PKID_Extract',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # windowed app – no console flash
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # add an .ico path here if you want a custom icon
)
