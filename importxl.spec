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
    'sqlite3.dbapi2',
    'sqlite3.dump',
    'sqlite3.user',
    'sqlite3.handler',
    'sqlite3.memory',
    'sqlite3.prepare_protocol',
    'sqlite3.connect',
    'sqlite3.complete_statement',
    'sqlite3.enable_callback_tracebacks',
    'sqlite3.register_adapter',
    'sqlite3.register_converter',
    'sqlite3.adapt',
    'sqlite3.converters',
    'sqlite3.paramstyle',
    'sqlite3.threadsafety',
    'sqlite3.apilevel',
    'sqlite3.version',
    'sqlite3.version_info',
    'sqlite3.sqlite_version',
    'sqlite3.sqlite_version_info',
    'sqlite3.OperationalError',
    'sqlite3.ProgrammingError',
    'sqlite3.IntegrityError',
    'sqlite3.DataError',
    'sqlite3.NotSupportedError',
    'sqlite3.Warning',
    'sqlite3.Error',
    'sqlite3.InterfaceError',
    'sqlite3.DatabaseError',
    'sqlite3.InternalError',
    'sqlite3.ConstraintError',
    'sqlite3.OperationalError',
    'sqlite3.ProgrammingError',
    'sqlite3.IntegrityError',
    'sqlite3.DataError',
    'sqlite3.NotSupportedError',
    'sqlite3.Warning',
    'sqlite3.Error',
    'sqlite3.InterfaceError',
    'sqlite3.DatabaseError',
    'sqlite3.InternalError',
    'sqlite3.ConstraintError',
])

# Collect data files
datas = [
    ('static', 'static'),
    ('config', 'config'),
]

# Create Data directory if it doesn't exist
import os
if not os.path.exists('Data'):
    os.makedirs('Data')
    with open('Data/README.txt', 'w') as f:
        f.write('This directory is for Excel files\n')
datas.append(('Data', 'Data'))

# Create config directory if it doesn't exist
if not os.path.exists('config'):
    os.makedirs('config')
    with open('config/README.md', 'w') as f:
        f.write('# Configuration Directory\n\nThis directory stores user preferences.\n')

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
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    icon=None,  # You can add an icon file here if you have one
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