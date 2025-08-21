# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    ['health-check-v12.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('temp_watermark1.png', '.'),
        ('arialbd.ttf', '.'),
        ('dp-splash.png', '.'),
        ('dplogo.png', '.'),
        ('dpoint1.png', '.'),
        ('blur.png', '.'),
        ('spcu', 'spcu'),
        ('serverinfo.txt', '.') 
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DPHealthCheck',
    debug=False,
    bootloader_ignore_signals=True,
    strip=True,
    upx=False, 
    upx_exclude=[],
    runtime_tmpdir='%TEMP%\\dphealth-tmp',
    console=False,
    disable_windowed_traceback=True,
    icon='dphealthcheck.ico',
    version='version_info.txt',
    
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=True,
    upx=False,
    upx_exclude=[],
    name='DPHealthCheck'
)
