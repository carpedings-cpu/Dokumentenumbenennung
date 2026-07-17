# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller-Spec fuer KPC Mail-Frage (eine Windows-.exe, onefile).

Bauen:  pyinstaller --noconfirm mail-frage.spec
Ergebnis: dist/KPC-Mail-Frage.exe
"""
import os

datas = []
binaries = []
# pywin32 fuer Outlook-COM (Postfach lesen).
hiddenimports = [
    "win32com", "win32com.client", "win32timezone", "pythoncom", "pywintypes",
    "briefing",
]

a = Analysis(
    [os.path.join("tool", "briefing", "mail_frage.py")],
    pathex=["tool/briefing"],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    a.binaries,
    a.datas,
    [],
    name="KPC-Mail-Frage",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # GUI: kein Konsolenfenster
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
