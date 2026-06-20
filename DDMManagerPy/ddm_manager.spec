# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec — DDM Manager v5.0 CustomTkinter
# Otimizado: single file, compressão, sem console

import os
from pathlib import Path

# Localiza os assets do CustomTkinter (temas e fontes)
import customtkinter
CTK_PATH = Path(customtkinter.__file__).parent

block_cipher = None

a = Analysis(
    ['ddm_manager.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Inclui tema e fontes do CustomTkinter
        (str(CTK_PATH / 'assets'), 'customtkinter/assets'),
    ],
    hiddenimports=[
        'customtkinter',
        'PIL', 'PIL.Image', 'PIL.ImageTk',   # dependência do CTk
        'docx', 'docx.oxml', 'docx.oxml.ns',
        'lxml', 'lxml.etree', 'lxml._elementpath',
        'win32com', 'win32com.client',
        'win32api', 'win32con', 'pywintypes',
        'openpyxl', 'openpyxl.styles',
        'winreg',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclui módulos pesados não usados → .exe menor e mais rápido
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'tkinter.test', 'unittest', 'email',
        'xml.etree.ElementTree',  # usa lxml
        'http', 'urllib', 'xmlrpc',
        'distutils', 'setuptools',
        'IPython', 'jupyter',
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
    upx=True,           # compressão UPX → .exe menor
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # sem janela de terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Ícone (opcional — descomente e coloque o caminho se tiver um .ico)
    # icon='icon.ico',
)
