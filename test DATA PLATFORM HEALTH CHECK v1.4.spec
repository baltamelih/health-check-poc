# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['health-check-v12.py'],
    pathex=[],
    binaries=[],
    datas=[('arialbd.ttf', '.'), ('dp-splash.png', '.'), ('dplogo.png', '.'), ('dpoint1.png', '.'), ('blur.png', '.'), ('serverinfo.txt', '.')],
    hiddenimports=[],
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
    name='test DATA PLATFORM HEALTH CHECK v1.4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='test DATA PLATFORM HEALTH CHECK v1.4',
)
