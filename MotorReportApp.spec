# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules
import os
try:
    from src._version import VERSION
except Exception:
    VERSION = "3.1.0"

datas = []
binaries = []
hiddenimports = [
    'src.config',
    'src.core',
    'src.data',
    'src.reports',
    'src.reports.builders',
    'src.analysis',
    'src.ui',
    'src.ui.core',
    'src.ui.core.event_handlers',
    'src.ui.core.state_manager',
    'src.ui.core.workflow_manager',
    'src.ui.core.status_manager',
    'src.ui.core.search_manager',
    'src.ui.core.report_manager',
    'src.ui.core.search_controller',
    'src.ui.core.selection_controller',
    'src.ui.core.report_generation_controller',
    'src.ui.core.file_picker_controller',
    'src.ui.core.configuration_controller',
    'src.ui.tabs',
    'src.ui.components',
    'src.ui.dialogs',
    'src.ui.utils',
    'src.services',
    'src.services.lf_registry_reader',
    'src.services.directory_locator',
    'src.services.noise_registry_reader',
    'src.services.noise_registry_loader',
    'src.services.noise_directory_cache',
    'src.services.registry_service',
    'src.services.test_lab_summary_loader',
    'src.services.lf_indexer',
    'src.utils',
    'src.validators',
]

# Collect all Flet and Flet web packages
tmp_ret = collect_all('flet')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('flet_web')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Collect additional Flet-related packages
tmp_ret = collect_all('flet_core')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('flet_runtime')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Collect web server dependencies
hiddenimports += collect_submodules('uvicorn')
hiddenimports += collect_submodules('starlette')
hiddenimports += collect_submodules('fastapi')
hiddenimports += collect_submodules('websockets')

# Collect large data-processing libs to ensure submodules are included
for pkg in ('pandas', 'numpy', 'openpyxl', 'xlsxwriter', 'PIL'):
    try:
        hiddenimports += collect_submodules(pkg)
    except Exception:
        # Skip optional or missing extensions (e.g., numba for pandas) during spec evaluation
        pass

# Add assets: include entire assets folder so runtime files are available
assets_root = os.path.join(os.getcwd(), 'assets')
if os.path.isdir(assets_root):
    for root, dirs, files in os.walk(assets_root):
        for fname in files:
            src_path = os.path.join(root, fname)
            rel_dir = os.path.relpath(root, assets_root)
            dest_dir = os.path.join('assets', rel_dir) if rel_dir != '.' else 'assets'
            datas.append((src_path, dest_dir))

# Defensive: do not bundle repository-level directory cache files
datas = [d for d in datas if not os.path.basename(d[0]).lower().startswith('directory_cache')]


a = Analysis(
    ['src\\main.py'],
    pathex=[
        'src',
        'src\\config',
        'src\\core',
        'src\\data',
        'src\\reports',
        'src\\reports\\builders',
        'src\\analysis',
        'src\\ui',
        'src\\ui\\core',
        'src\\ui\\tabs',
        'src\\ui\\components',
        'src\\ui\\dialogs',
        'src\\ui\\utils',
        'src\\services',
        'src\\utils',
        'src\\validators',
    ],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
