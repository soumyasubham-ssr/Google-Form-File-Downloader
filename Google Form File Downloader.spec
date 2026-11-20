# -*- mode: python ; coding: utf-8 -*-
# Build:  pyinstaller --noconfirm "Google Form File Downloader.spec"
import re
from pathlib import Path

from PyInstaller.utils.win32.versioninfo import (FixedFileInfo, StringFileInfo, StringStruct,
                                                 StringTable, VarFileInfo, VarStruct,
                                                 VSVersionInfo)

VERSION = re.search(r'__version__ = "([^"]+)"', Path("gffd/__init__.py").read_text()).group(1)
NUMS = tuple(int(n) for n in (VERSION.split(".") + ["0"] * 4)[:4])

version_info = VSVersionInfo(
    ffi=FixedFileInfo(filevers=NUMS, prodvers=NUMS),
    kids=[
        StringFileInfo([StringTable("040904B0", [
            StringStruct("CompanyName", "Soumya Subham Rout"),
            StringStruct("FileDescription", "Google Form File Downloader"),
            StringStruct("FileVersion", VERSION),
            StringStruct("InternalName", "GoogleFormFileDownloader"),
            StringStruct("LegalCopyright", "© 2026 Soumya Subham Rout"),
            StringStruct("OriginalFilename", "Google Form File Downloader.exe"),
            StringStruct("ProductName", "Google Form File Downloader"),
            StringStruct("ProductVersion", VERSION),
        ])]),
        VarFileInfo([VarStruct("Translation", [1033, 1200])]),
    ],
)

a = Analysis(
    ["google_form_file_downloader.py"],
    datas=[("assets/template.xlsx", "assets"), ("assets/icon.ico", "assets")],
    excludes=["pandas", "numpy", "PIL", "pytest", "unittest", "pydoc", "test", "lib2to3"],
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Google Form File Downloader",
    debug=False,
    strip=False,
    upx=True,
    console=False,
    icon="assets/icon.ico",
    version=version_info,
)
