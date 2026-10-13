# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/app.py'],
    pathex=[],
    binaries=[],
    datas=[('E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/assets', 'assets')],
    hiddenimports=['docxtpl', 'docx', 'fitz', 'pymupdf', 'pypdf', 'pandas', 'openpyxl', 'core.updater', 'win32com', 'win32com.client'],
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
    [],
    exclude_binaries=True,
    name='Canido Doc Manager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='C:/Users/User/AppData/Local/Temp/9f550974-d132-4dfa-8a34-52dc1f103756',
    icon=['E:/@Medi/@002 Gemini Projects/@0001_canido_doc_manager/assets/logo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Canido Doc Manager',
)
