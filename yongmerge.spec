# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None


def _resource(name):
    return os.path.join(os.getcwd(), name)


hidden_imports = collect_submodules("win32com")

a = Analysis(
    ["main_app.py"],
    pathex=[],
    binaries=[],
    datas=[
        (_resource("PretendardVariable.ttf"), "."),
        (_resource("yongmerge.ico"), "."),
        (_resource("yongpdf_donation.jpg"), "."),
        (_resource("YongMerge_img.png"), "."),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="YongMerge",
    icon=_resource("yongmerge.ico"),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
)
