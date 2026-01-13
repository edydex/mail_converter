# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Mayo's Mail Converter

To build:
    pyinstaller mail_converter.spec

For Windows:
    pyinstaller mail_converter.spec --onefile

For macOS app bundle:
    pyinstaller mail_converter.spec
"""

import sys
import os
from pathlib import Path

block_cipher = None

# Get the project root
PROJECT_ROOT = Path(SPECPATH)
BUILD_DIR = PROJECT_ROOT / 'build'

# Collect binary files for Windows
binaries_list = []
datas_list = []

# Add readpst.exe and libpst DLL for Windows
if sys.platform == 'win32':
    readpst_path = BUILD_DIR / 'bin' / 'readpst.exe'
    libpst_path = BUILD_DIR / 'bin' / 'libpst-4.dll'
    
    if readpst_path.exists():
        binaries_list.append((str(readpst_path), '.'))
    if libpst_path.exists():
        binaries_list.append((str(libpst_path), '.'))
    
    # Add poppler binaries
    poppler_bin = BUILD_DIR / 'poppler' / 'poppler-24.08.0' / 'Library' / 'bin'
    if poppler_bin.exists():
        for dll in poppler_bin.glob('*.dll'):
            binaries_list.append((str(dll), 'poppler'))
        for exe in poppler_bin.glob('*.exe'):
            binaries_list.append((str(exe), 'poppler'))

# Add assets if they exist
if (PROJECT_ROOT / 'assets').exists():
    datas_list.append(('assets', 'assets'))

# Collect all source files
a = Analysis(
    ['main.py'],
    pathex=[str(PROJECT_ROOT)],
    binaries=binaries_list,
    datas=datas_list,
    hiddenimports=[
        'PIL._tkinter_finder',
        'reportlab.graphics.barcode.common',
        'reportlab.graphics.barcode.code128',
        'reportlab.graphics.barcode.code93',
        'reportlab.graphics.barcode.code39',
        'reportlab.graphics.barcode.usps',
        'reportlab.graphics.barcode.usps4s',
        'reportlab.graphics.barcode.ecc200datamatrix',
        'pkg_resources.py2_warn',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'numpy.distutils',
        'IPython',
        'jupyter',
    ],
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
    name='MayosMailConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico' if sys.platform == 'win32' and (PROJECT_ROOT / 'assets' / 'icon.ico').exists() else None,
)

# macOS app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name="Mayo's Mail Converter.app",
        icon='assets/icon.icns' if (PROJECT_ROOT / 'assets' / 'icon.icns').exists() else None,
        bundle_identifier='com.edydex.mayosmailconverter',
        info_plist={
            'NSHighResolutionCapable': 'True',
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
        },
    )
