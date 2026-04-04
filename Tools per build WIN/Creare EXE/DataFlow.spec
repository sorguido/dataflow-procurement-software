# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for DataFlow Procurement Software
Builds Windows executable with proper Tcl/Tk support for Tkinter.

Usage:
    pyinstaller dataflow.spec

This will create:
    - dist/dataflow/ (folder containing dataflow.exe and dependencies)

Note: This uses one-folder mode (not one-file) for better Tcl/Tk compatibility.
"""

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# ============================================================================
# AUTOMATIC TCL/TK DETECTION (Windows compatibility fix)
# ============================================================================
# This solves "Can't find a usable init.tcl" error on Windows
# By explicitly including Tcl/Tk runtime files required by Tkinter

added_files = []

# Add project resources
added_files.extend([
    ('add_data', 'add_data'),
    ('locale', 'locale'),
])

# ============================================================================
# TCL/TK RUNTIME FILES (Critical for Tkinter on Windows)
# ============================================================================
# PyInstaller sometimes fails to auto-detect Tcl/Tk files.
# We add them explicitly using Python's tkinter module to find the paths.

if sys.platform == 'win32':
    import tkinter
    import _tkinter
    
    # Get the Tcl/Tk library directory from tkinter
    # This is portable and doesn't use hardcoded paths
    tkinter_dir = os.path.dirname(tkinter.__file__)
    tcl_dir = os.path.join(tkinter_dir, 'tcl')
    
    # If tcl directory exists in tkinter installation, add it
    if os.path.exists(tcl_dir):
        # Find tcl8.6 and tk8.6 directories
        for item in os.listdir(tcl_dir):
            item_path = os.path.join(tcl_dir, item)
            if os.path.isdir(item_path):
                # Add tcl8.6, tk8.6, and other tcl subdirectories
                if item.startswith('tcl') or item.startswith('tk'):
                    added_files.append((item_path, os.path.join('tcl', item)))
        
        print(f"✓ Tcl/Tk libraries added from: {tcl_dir}")
    else:
        # Fallback: try to find Tcl/Tk in Python installation directory
        # This works for standard Python installations
        python_dir = os.path.dirname(sys.executable)
        tcl_fallback = os.path.join(python_dir, 'tcl')
        
        if os.path.exists(tcl_fallback):
            for item in os.listdir(tcl_fallback):
                item_path = os.path.join(tcl_fallback, item)
                if os.path.isdir(item_path):
                    if item.startswith('tcl') or item.startswith('tk'):
                        added_files.append((item_path, os.path.join('tcl', item)))
            print(f"✓ Tcl/Tk libraries added from Python installation: {tcl_fallback}")
        else:
            print("⚠ Warning: Could not auto-detect Tcl/Tk directories.")
            print("  Tkinter may not work in the built executable.")
            print(f"  Searched in: {tcl_dir} and {tcl_fallback}")

# ============================================================================
# ANALYSIS - MODULE AND DEPENDENCY DETECTION
# ============================================================================

a = Analysis(
    ['dataflow.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        # Explicitly include Tkinter and Tcl/Tk internals
        'tkinter',
        '_tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.font',
        
        # Custom application modules
        'database',
        'database.db_helpers',
        'services',
        'services.app_paths',
        'services.startup_service',
        'ui',
        'ui.help_window',
        'ui.dialogs',
        'ui.dialogs.common_dialogs',
        'ui.windows',
        'ui.windows.view_request_window',
        'ui.windows.edit_suppliers_window',
        'ui.windows.edit_reference_window',
        'ui.windows.notes_window',
        'ui.windows.purchase_order_window',
        'ui.windows.attachment_window',
        'ui.windows.sqdc_analysis_window',
        'utils',
        'utils.format_utils',
        'utils.i18n_utils',
        'utils.resource_utils',
        'utils.string_utils',
        'utils.user_utils',
        'utils.validation_utils',
        'utils.window_utils',
        
        # Required third-party packages
        'PIL',
        'PIL._tkinter_finder',
        'PIL.Image',
        'PIL.ImageTk',
        'openpyxl',
        'openpyxl.styles',
        'tksheet',
        'tkcalendar',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude unused modules to reduce size
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'IPython',
        'jupyter',
        'pytest',
        'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ============================================================================
# PYZ - Python Archive (compiled .pyc files)
# ============================================================================

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ============================================================================
# EXE - Executable Configuration (ONE-FOLDER MODE)
# ============================================================================
# Using one-folder mode for better Tcl/Tk compatibility
# The executable will be in dist/dataflow/dataflow.exe
# All dependencies (including Tcl/Tk) will be in dist/dataflow/ folder

exe = EXE(
    pyz,
    a.scripts,
    [],  # Empty list = one-folder mode
    exclude_binaries=True,  # Important: keep binaries separate for one-folder mode
    name='dataflow',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to False for GUI application (no console window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='add_data/DataFlow.ico',
    manifest='app.manifest.xml',
)

# ============================================================================
# COLLECT - Bundle Everything (ONE-FOLDER MODE)
# ============================================================================
# This collects all binaries, zipfiles, and data into the dist folder

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='dataflow',
)

# ============================================================================
# BUILD VERIFICATION
# ============================================================================

print("\n" + "="*70)
print("PyInstaller Build Configuration Summary")
print("="*70)
print(f"Target platform: {sys.platform}")
print(f"Build mode: ONE-FOLDER")
print(f"Output location: dist/dataflow/")
print(f"Executable: dist/dataflow/dataflow.exe")
print(f"Resources included: {len(added_files)} directories/files")
print(f"Hidden imports: {len(a.hiddenimports)} modules")
print("="*70)
print("\nAfter build, test with:")
print("  1. cd dist/dataflow")
print("  2. dataflow.exe")
print("="*70 + "\n")