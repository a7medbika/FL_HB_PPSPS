# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['D:\\P R O J E C T S\\IT\\FL_HB_PPSPS\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('D:/P R O J E C T S/IT/FL_HB_PPSPS/icons', 'icons/'), ('D:/P R O J E C T S/IT/FL_HB_PPSPS/UI', 'UI/'), ('D:/P R O J E C T S/IT/FL_HB_PPSPS/staticPDFs', 'staticPDFs/'), ('D:/P R O J E C T S/IT/FL_HB_PPSPS/Base.docx', '.'), ('D:\\P R O J E C T S\\IT\\FL_HB_PPSPS\\key', '.'), ('D:\\P R O J E C T S\\IT\\FL_HB_PPSPS\\ppsps.mp4', '.')],
    hiddenimports=[],
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
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
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
    icon='D:\\P R O J E C T S\\IT\\FL_HB_PPSPS\\icons\\icon (1).ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
