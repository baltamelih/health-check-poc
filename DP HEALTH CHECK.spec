# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[('error_log.txt', '.'), ('dplogo.png', '.'), ('dpoint.png', '.'), ('arialbd.ttf', '.'), ('blur.png', '.')],
    datas=[('encoded_script.py', '.'), ('encoded_script2.py', '.'), ('backuplogrunning.py', '.'), ('policynotcheckedlogin.py', '.'), ('samepasswordlogins.py', '.'), ('drop.py', '.'), ('new_script.py', '.'), ('xpcmdshell.py', '.')],
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
    name='DP HEALTH CHECK',
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
    name='DP HEALTH CHECK',
)
