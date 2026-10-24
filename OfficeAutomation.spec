# -*- mode: python ; coding: utf-8 -*-
import sys
import os

block_cipher = None

added_files = [
    ('templates', 'templates'),
]

hidden_imports = [
    'openpyxl',
    'openpyxl.cell._writer',
    'xlrd',
    'xlwt',
    'xlutils',
    'xlutils.copy',
    'pdfplumber',
    'pdfminer',
    'pdfminer.six',
    'pdfminer.high_level',
    'pypdfium2',
    'pypdf',
    'flask',
    'jinja2',
    'werkzeug',
    'webview',
    'clr',
    'clr_loader',
    'pythonnet',
    'tkinter',
    'tkinter.filedialog',
]

a = Analysis(
    ['main_desktop.py'],
    pathex=['.'],
    binaries=[],
    datas=added_files,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'numpy'],
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
    name='OfficeAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico'
)
