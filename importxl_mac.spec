# -*- mode: python ; coding: utf-8 -*-
import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# Collect all necessary modules
hiddenimports = collect_submodules('pandas')
hiddenimports.extend(collect_submodules('openpyxl'))
hiddenimports.extend(collect_submodules('xlsxwriter'))
hiddenimports.extend(collect_submodules('sqlite3'))
hiddenimports.extend(collect_submodules('flask'))
hiddenimports.extend(collect_submodules('werkzeug'))
hiddenimports.extend(collect_submodules('jinja2'))
hiddenimports.extend(collect_submodules('markupsafe'))
hiddenimports.extend(collect_submodules('itsdangerous'))
hiddenimports.extend(collect_submodules('click'))
hiddenimports.extend(collect_submodules('blinker'))

# Add specific modules that might be missed
hiddenimports.extend([
    'pandas.io.excel._xlsxwriter',
    'pandas.io.excel._openpyxl',
    'openpyxl.cell',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'xlsxwriter.workbook',
    'xlsxwriter.worksheet',
    'flask.json',
    'flask.sessions',
    'werkzeug.middleware',
    'werkzeug.security',
    'jinja2.ext',
    'markupsafe._markupsafe',
])

# Collect data files
datas = [
    ('static', 'static'),
    ('Data', 'Data'),
    ('config', 'config'),
]

# Add any additional data files from packages
datas.extend(collect_data_files('pandas'))
datas.extend(collect_data_files('openpyxl'))
datas.extend(collect_data_files('xlsxwriter'))

a = Analysis([
    'importxl_web.py',
],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='LAB14_Excel_Importer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='LAB14_Excel_Importer'
) 