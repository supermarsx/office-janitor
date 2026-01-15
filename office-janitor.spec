# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

# Collect all OEM files (XML configs and binaries)
oem_dir = Path('oem')
oem_files = []
if oem_dir.exists():
    for f in oem_dir.iterdir():
        if f.is_file():
            oem_files.append((str(f), 'oem'))

# Collect OfficeScrubber VBS scripts
offscrub_dir = Path('OfficeScrubber/bin')
offscrub_files = []
if offscrub_dir.exists():
    for f in offscrub_dir.iterdir():
        if f.is_file():
            offscrub_files.append((str(f), 'OfficeScrubber/bin'))

# Include VERSION file from package
version_file = Path('src/office_janitor/VERSION')
datas_list = oem_files + offscrub_files
if version_file.exists():
    datas_list.append((str(version_file), 'office_janitor'))

a = Analysis(
    ['oj_entry.py'],
    pathex=['src'],
    binaries=[],
    datas=datas_list,
    hiddenimports=['office_janitor', 'office_janitor.main'],
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
    name='office-janitor',
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
    uac_admin=True,
)
