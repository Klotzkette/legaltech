# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller build spec for Tom's Super Simple Word-Gliederungs-Retter.

macOS .app bundle.
Run with:  pyinstaller build_mac.spec --noconfirm
"""

import os
import sys

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=[os.path.abspath('src')],
    binaries=[],
    datas=[
        ('src/gui.py', 'src'),
        ('src/word_processor.py', 'src'),
        ('src/ai_engine.py', 'src'),
    ],
    hiddenimports=[
        'docx',
        'docx.opc',
        'docx.opc.part',
        'docx.opc.pkgwriter',
        'docx.opc.constants',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.oxml.parser',
        'docx.oxml.shared',
        'docx.oxml.text.paragraph',
        'docx.oxml.text.run',
        'docx.oxml.numbering',
        'docx.parts',
        'docx.parts.numbering',
        'docx.parts.document',
        'docx.parts.styles',
        'docx.enum',
        'docx.enum.style',
        'docx.shared',
        'docx.styles',
        'docx.styles.style',
        'docx.text',
        'docx.text.paragraph',
        'docx.text.run',
        'lxml',
        'lxml.etree',
        'lxml._elementpath',
        'openai',
        'PyQt6',
        'PyQt6.QtWidgets',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.sip',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'scipy', 'pandas'],
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
    name='Word-Gliederungs-Retter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

app = BUNDLE(
    exe,
    name='Word-Gliederungs-Retter.app',
    icon=None,
    bundle_identifier='com.toms.wordgliederungsretter',
    info_plist={
        'CFBundleDisplayName': "Word-Gliederungs-Retter",
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '11.0',
        'NSRequiresAquaSystemAppearance': False,
    },
)
