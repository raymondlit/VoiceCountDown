# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['count_down12.py'],
    pathex=[],
    binaries=[],
    datas=[('comtypes\\gen', 'comtypes\\gen')],
    hiddenimports=['pyttsx3', 'pyttsx3.drivers.sapi5', 'pyttsx3.voice', 'comtypes.client', 'comtypes.gen', 'win32gui', 'win32process', 'win32api', 'ctypes', 'numpy', 'psutil'],
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
    name='count_down12',
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
    uac_admin=True,
    icon=['timer.ico'],
)
