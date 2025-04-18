# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[('encoded_script.py', '.'), ('encoded_script2.py', '.'), ('encoded_script.py', '.'), ('backuplogrunning.py', '.'), ('policynotcheckedlogin.py', '.'), ('samepasswordlogins.py', '.'), ('new_script.py', '.'), ('xpcmdshell.py', '.'), ('drop.py', '.'), ('spcu', 'spcu.'), ('arialbd.ttf', '.'), ('dp-splash.png', '.'), ('dplogo.png', '.'), ('dpoint.png', '.'), ('blur.png', '.'), ('vlfcount.py', '.')],
    datas=[],
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
    name='DATA PLATFORM HEALTH CHECK',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DATA PLATFORM HEALTH CHECK',
)
