# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\sebas\\Downloads\\OutlookArchiver\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\sebas\\Downloads\\OutlookArchiver\\config.py', '.'), ('C:\\Users\\sebas\\Downloads\\OutlookArchiver\\archiver.py', '.'), ('C:\\Users\\sebas\\Downloads\\OutlookArchiver\\scheduler.py', '.'), ('C:\\Users\\sebas\\Downloads\\OutlookArchiver\\logger.py', '.'), ('C:\\Users\\sebas\\Downloads\\OutlookArchiver\\gui.py', '.'), ('C:\\Users\\sebas\\Downloads\\OutlookArchiver\\startup.py', '.'), ('C:\\Users\\sebas\\Downloads\\OutlookArchiver\\wizard.py', '.')],
    hiddenimports=['win32com.client', 'win32com.shell', 'pywintypes', 'winreg'],
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
    name='OutlookArchiver',
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
)
