# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['mainGUI.py'],
    pathex=[],
    binaries=[],
    datas=[('efl3pl.png', '.'), ('chromedriver.exe', '.'), ('icon.ico', '.'), ('C:/Users/96598/AppData/Local/Programs/Python/Python310/Lib/site-packages/customtkinter', 'customtkinter/')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'hook', 'PIL', 'pandas', 'opencv', 'beautifulsoup4', 'pillow', 'tornado', 'plotly', 'sqlite3', 'kaleido'],
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
    name='GUI',
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='OBSS',
)
