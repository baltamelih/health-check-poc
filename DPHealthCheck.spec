# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    ['health-check-v12.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('public/temp_watermark1.png', '.'),
        ('public/arialbd.ttf', '.'),
        ('public/dp-splash.png', '.'),
        ('public/dplogo.png', '.'),
        ('public/dpoint1.png', '.'),
        ('public/blur.png', '.'),
        ('spcu', 'spcu'),
        ('src/serverinfo.txt', '.') ,
        ('src', 'src'),
        ('.env', '.')
        
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
