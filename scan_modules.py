# -*- mode: python ; coding: utf-8 -*-
import os, sys
from PyInstaller.utils.hooks import collect_dynamic_libs, collect_data_files

block_cipher = None

# --- Matplotlib: sadece Agg backend ---
# Kod tarafında: matplotlib.use("Agg") satırı *pyplot importundan önce* kalmalı.

# --- PyQt5: yalın plugin seti ---
# Sadece gerekli Qt pluginlerini toplayalım: 'platforms/qwindows', 'imageformats/{qjpeg,qpng,qico}'
qt_dynlibs = []
qt_datas   = []

# PyInstaller default'u tüm pluginleri sürükler. Biz hedefli toplayacağız:
# PyQt5 runtime kütüphaneleri (Qt5*.dll) yine otomatik gelir; burada *plugin* DLL’lerini daraltıyoruz.
qt_plugins_base = os.path.join('PyQt5', 'Qt5', 'plugins')

needed_plugins = [
    os.path.join(qt_plugins_base, 'platforms', 'qwindows.dll'),
    os.path.join(qt_plugins_base, 'imageformats', 'qjpeg.dll'),
    os.path.join(qt_plugins_base, 'imageformats', 'qpng.dll'),
    os.path.join(qt_plugins_base, 'imageformats', 'qico.dll'),
]

# Bu dosyaları bulunduğu yerden alıp 'PyQt5/Qt5/plugins/...'
# altında aynı göreli yolla dist’e kopyalayalım:
for p in needed_plugins:
    # PyInstaller, bu dosyaları kurulu ortamında bulabilsin diye collect_dynamic_libs kullan:
    base_mod = 'PyQt5'
    for src, dest in collect_dynamic_libs(base_mod):
        if src.replace('\\', '/').endswith(p.replace('\\', '/')):
            rel_dir = os.path.dirname(p)
            qt_dynlibs.append((src, rel_dir))
            break

# (Ekstra: bazı sistemlerde png/jpeg pluginleri data olarak paketlenebilir)
qt_datas += [(src, os.path.dirname(p)) for (src, p) in qt_dynlibs]

# Matplotlib’in devasa data setini almamak için data toplamıyoruz.
# Eğer özel fontları sen veriyorsan (arialbd.ttf gibi), yeterli.

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
        ('serverinfo.txt', '.'),
        # PyQt pluginleri hedefli ekle:
        *qt_datas,
    ],
    hiddenimports=[
        # Matplotlib sadece Agg:
        'matplotlib.backends.backend_agg',
        'matplotlib.pyplot',
        # Pandas Excel motoru:
        'openpyxl',
        # PyQt5 çekirdek modüller bazen dinamik çözülür:
        'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets',
    ],
    excludes=[
        # GUI stack dışı:
        'tkinter', 'tcl',
        'PyQt6', 'PySide2', 'PySide6',

        # Matplotlib şişman kısımlar:
        'matplotlib.tests',
        'matplotlib.backends.backend_qt5',
        'matplotlib.backends.backend_qt4',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.backends.backend_wx',
        'matplotlib.backends._macosx',

        # Pandas/Numpy test ve araçları:
        'pandas.tests', 'pandas.plotting', 'pandas.io.clipboards',
        'numpy.f2py', 'numpy.testing', 'numpy.tests', 'numpy.matrixlib',

        # PyQt5’te gereksiz plugin aileleri:
        'PyQt5.QtWebEngineWidgets', 'PyQt5.QtWebEngineCore',
        'PyQt5.QtWebChannel', 'PyQt5.QtMultimedia', 'PyQt5.QtMultimediaWidgets',
        'PyQt5.QtNetwork', 'PyQt5.QtBluetooth', 'PyQt5.QtLocation',
        'PyQt5.QtPositioning', 'PyQt5.QtSensors', 'PyQt5.QtSerialPort',
        'PyQt5.QtSql', 'PyQt5.QtWebSockets', 'PyQt5.QtNfc',
        # (kullanmıyorsan ekleyebilirsin; koduna göre daralt)
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=True,  # şeffaflık: AV için daha iyi sinyal
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='DPHealthCheck',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,      # AV dostu
    upx=False,        # AV dostu
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=True,
    icon='dphealthcheck.ico',
    version='version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries + qt_dynlibs,  # hedefli Qt plugin DLL’lerini ekle
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='DPHealthCheck'
)
