# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
import os


def collect_folder(folder_name):
    items = []
    for root, _, files in os.walk(folder_name):
        for filename in files:
            src = os.path.join(root, filename)
            rel_dir = os.path.relpath(root, ".")
            items.append((src, rel_dir))
    return items


datas = []
datas += collect_data_files("babel")
datas += collect_folder("add_data")
datas += collect_folder("locale")

a = Analysis(
    ['dataflow.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
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
    icon='add_data\\DataFlow.ico',
    manifest='app.manifest.xml',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='dataflow',
)