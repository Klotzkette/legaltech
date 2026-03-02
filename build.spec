# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller build spec for Tom's Super Simple Word-Gliederungs-Retter.
Run with:  pyinstaller build.spec --noconfirm
"""

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=[os.path.abspath('src')],
    binaries=[],
    datas=[
        ('src', 'src'),
    ],
    hiddenimports=[
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'lxml',
        'lxml.etree',
        'openai',
        'PyQt6',
        'PyQt6.QtWidgets',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
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
    [],
    exclude_binaries=True,
    name='Word-Gliederungs-Retter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Word-Gliederungs-Retter',
)
