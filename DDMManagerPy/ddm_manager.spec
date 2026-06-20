# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — DDM Manager v5.0 CustomTkinter

import sys
from pathlib import Path

import customtkinter
CTK_PATH = Path(customtkinter.__file__).parent

block_cipher = None

a = Analysis(
    ['ddm_manager.py'],
    pathex=[],
    binaries=[],
    datas=[
        (str(CTK_PATH / 'assets'), 'customtkinter/assets'),
    ],
    hiddenimports=[
        'customtkinter',
        'PIL', 'PIL.Image', 'PIL.ImageTk',
        'docx', 'docx.oxml', 'docx.oxml.ns',
        'lxml', 'lxml.etree', 'lxml._elementpath', 'lxml.builder',
        'win32com', 'win32com.client',
        'win32api', 'win32con', 'pywintypes',
        'openpyxl', 'openpyxl.styles',
        'winreg',
        'tkinter', 'tkinter.filedialog', 'tkinter.messagebox',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'IPython', 'jupyter', 'notebook',
        'distutils', 'setuptools', 'pkg_resources',
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
    name='DDM_Manager_Metalfrio',
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
)
