# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules
import os
import sys

# --- Configuration ---
PROJECT_ROOT = os.getcwd()
SRC_DIR = os.path.join(PROJECT_ROOT, 'src')

# Ensure we can import from src during build to check version
sys.path.insert(0, PROJECT_ROOT)

try:
    from src._version import VERSION
except ImportError:
    VERSION = "0.0.0"

# --- Collection ---
datas = []
binaries = []
hiddenimports = []

# 1. Collect all submodules from 'src' automatically
# This replaces the long manual list and ensures new files are always included.
try:
    hiddenimports += collect_submodules('src')
except Exception as e:
    print(f"WARNING: Failed to collect src submodules: {e}")

# FORCE INCLUDE CRITICAL EXCEL LIBRARIES
# PyInstaller misses these because they are often loaded dynamically by pandas (engine='xlrd')
hiddenimports += [
    'xlrd', 
    'openpyxl', 
    'openpyxl.cell._writer',
    'pandas',
    'numpy'
]

# 2. Collect Flet and Web dependencies (Crucial for UI)
# collect_all returns (datas, binaries, hiddenimports)
for package in ['flet', 'flet_web', 'flet_core', 'flet_runtime']:
    try:
        tmp_ret = collect_all(package)
        datas += tmp_ret[0]
        binaries += tmp_ret[1]
        hiddenimports += tmp_ret[2]
    except Exception:
        pass

# 3. Collect Data Processing Libraries
# These often have hidden C-extensions or lazy imports
for package in ['pandas', 'numpy', 'openpyxl', 'xlsxwriter', 'PIL', 'xlrd', 'uvicorn', 'fastapi', 'starlette', 'websockets']:
    try:
        hiddenimports += collect_submodules(package)
    except Exception:
        pass

# 4. Collect Assets
# We copy the whole assets folder to the root of the bundle
assets_path = os.path.join(PROJECT_ROOT, 'assets')
if os.path.exists(assets_path):
    datas.append((assets_path, 'assets'))

# --- Build Analysis ---
a = Analysis(
    ['src\\main.py'],
    pathex=[PROJECT_ROOT],  # ONLY Project Root to avoid import confusion
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['src/hooks/runtime_hook.py'],
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
    name=f'MotorReportApp-{VERSION}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets\\logo.ico'],
)
# Single-file build does not use COLLECT
# coll = COLLECT(...)
