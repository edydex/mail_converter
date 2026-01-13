# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Mail Converter

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

# Collect all source files
a = Analysis(
    ['main.py'],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=[
        # Include assets if they exist
        ('assets', 'assets') if (PROJECT_ROOT / 'assets').exists() else (None, None),
    ],
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

# Filter out None entries from datas
a.datas = [d for d in a.datas if d[0] is not None]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MailConverter',
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
