# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/app.py'],
    pathex=[],
    binaries=[],
    datas=[('E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/assets', 'assets')],
    hiddenimports=['docxtpl', 'docx', 'fitz', 'pymupdf', 'pypdf', 'pandas', 'openpyxl', 'win32com', 'win32com.client'],
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
    name='Canido Doc Manager',
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
    version='C:/Users/User/AppData/Local/Temp/f9fcc12d-32a2-46c9-affd-6fd49991abd5',
    icon=['E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/assets/logo.ico'],
)
