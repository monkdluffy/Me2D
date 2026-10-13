# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/assets', 'assets')]
binaries = []
hiddenimports = ['docxtpl', 'docx', 'fitz', 'pymupdf', 'pypdf', 'pandas', 'openpyxl', 'flet', 'flet_desktop', 'win32com', 'win32com.client']
tmp_ret = collect_all('flet_desktop')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('flet')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='SMAP_AutoDoc_Studio_Single',
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
    icon=['E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/assets/logo.ico'],
)
