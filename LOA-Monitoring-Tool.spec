# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

is_onedir = os.environ.get('LOA_BUILD_ONEDIR', '').strip().lower() in ('1', 'true', 'yes')
dataverse_config = os.environ.get('LOA_DATAVERSE_CONFIG', '').strip()

datas = [('web', 'web')]
if dataverse_config and os.path.exists(dataverse_config):
    datas.append((dataverse_config, '.'))
elif os.path.exists('dataverse_config.json'):
    datas.append(('dataverse_config.json', '.'))

binaries = []
hiddenimports = ['loa_engine', 'dataverse', 'app_paths', 'updater', 'openpyxl', 'requests', 'msal']

tmp_ret = collect_all('webview')
datas += tmp_ret[0]; binaries += tmp_ret[1]
# Windows only: skip the Android/Qt/GTK/Cocoa back-ends.
hiddenimports += [m for m in tmp_ret[2] if not any(p in m for p in ('android', '.qt', '.gtk', '.cocoa', '.cef'))]

tmp_ret = collect_all('velopack')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

tmp_ret = collect_all('msal')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Packages the app never uses. Building from a Python install that has them
# (torch, transformers, scipy, ...) pulled them in and made the exe ~400 MB,
# which is unpacked to %TEMP% on every launch. Build with build.ps1, which uses
# a clean virtual environment; this list is a second safety net.
EXCLUDES = [
    'torch', 'torchgen', 'torchvision', 'torchaudio', 'transformers', 'huggingface_hub', 'accelerate',
    'tensorflow', 'keras', 'scipy', 'sklearn', 'sympy', 'mpmath', 'networkx', 'pandas', 'numpy',
    'matplotlib', 'PIL', 'IPython', 'jupyter', 'notebook', 'pytest', 'sqlalchemy', 'psycopg', 'psycopg2',
    'redis', 'boto3', 'botocore', 's3transfer', 'azure', 'google', 'grpc', 'opentelemetry', 'pydantic',
    'aiohttp', 'uvicorn', 'fastapi', 'starlette', 'websockets', 'mcp', 'authlib', 'joserfc', 'narwhals',
    'xlsxwriter', 'pygments', 'rich', 'markdown_it', 'joblib', 'fsspec', 'tkinter',
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'qtpy', 'gi', 'cefpython3',
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

if is_onedir:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='LOA-Monitoring-Tool',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon='icon.ico' if os.path.exists('icon.ico') else None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name='LOA-Monitoring-Tool',
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name='LOA-Monitoring-Tool',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon='icon.ico' if os.path.exists('icon.ico') else None,
    )
