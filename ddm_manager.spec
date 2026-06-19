# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    ['ddm_manager.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'docx', 'docx.oxml', 'docx.oxml.ns',
        'lxml', 'lxml.etree', 'lxml._elementpath', 'lxml.builder',
        'win32com', 'win32com.client', 'win32api', 'win32con',
        'pywintypes',
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
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
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
