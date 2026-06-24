# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller-Spec fuer das KPC-Morgenbriefing (eine Windows-.exe, onefile).

Bauen:  pyinstaller --noconfirm briefing.spec
Ergebnis: dist/KPC-Morgenbriefing.exe
"""
import os

# Projektliste mitbuendeln (zur Gruppierung nach Projekt); optional.
datas = []
_mapping = os.path.join("tool", "triage", "projekte_mapping.json")
if os.path.exists(_mapping):
    datas.append((_mapping, "."))

binaries = []
# pywin32 fuer Outlook-COM (Lesen der Mails + Kalendertermine anlegen).
hiddenimports = [
    "win32com", "win32com.client", "win32timezone", "pythoncom", "pywintypes",
]

a = Analysis(
    [os.path.join("tool", "briefing", "briefing.py")],
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
    name="KPC-Morgenbriefing",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # GUI/Hintergrund: kein Konsolenfenster
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
