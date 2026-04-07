# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

datas = [
    ('locale', 'locale'),

    # Template Excel
    ('add_data', 'add_data'),

    # Asset grafici anche in root
    ('add_data/DataFlow.ico', '.'),
    ('add_data/Logo.png', '.'),
    ('add_data/Logo_44x44.png', '.'),
    ('add_data/Logo_50x50.png', '.'),
    ('add_data/Logo_150x150.png', '.'),
    ('add_data/logo_dataflow.png', '.'),
]

# Babel (fondamentale)
datas += collect_data_files('babel')

a = Analysis(
    ['dataflow.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'PIL._tkinter_finder',   # 🔴 FIX CRITICO
    ],
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
    name='dataflow',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='dataflow_appimage',
    contents_directory='.',   # 🔴 QUESTO deve funzionare
)
