# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller-Spec fuer den Dokumenten-Umbenenner (eine Windows-.exe, onefile).

Bauen:  pyinstaller --noconfirm dokumenten-umbenenner.spec
Ergebnis: dist/Dokumenten-Umbenenner.exe
"""
import os
from PyInstaller.utils.hooks import collect_all

# Die VA-Regeln (SKILL.md) werden in die .exe gebuendelt, damit der API-Modus
# den Systemprompt auch ohne Quelldateien findet. Die Projektliste
# (projekte_mapping.json) liefert Stichworte fuer die Projekt-Einsortierung.
datas = [(os.path.join(".claude", "skills", "dokumentenumbenennung", "SKILL.md"), ".")]
_mapping = os.path.join("tool", "triage", "projekte_mapping.json")
if os.path.exists(_mapping):
    datas.append((_mapping, "."))
binaries = []
hiddenimports = ["va_rules", "pdf_text", "api_client", "gemini_client",
                 "email_extract", "projekt_zuordnung"]

# Anthropic-SDK mitbuendeln, falls installiert (fuer den optionalen API-Modus).
# Fehlt es, bleibt der Offline-Modus voll funktionsfaehig.
try:
    a_datas, a_binaries, a_hidden = collect_all("anthropic")
    datas += a_datas
    binaries += a_binaries
    hiddenimports += a_hidden
except Exception:
    pass

# tkinterdnd2 (Drag & Drop) inkl. der tkdnd-Bibliotheksdateien mitbuendeln.
try:
    d_datas, d_binaries, d_hidden = collect_all("tkinterdnd2")
    datas += d_datas
    binaries += d_binaries
    hiddenimports += d_hidden
except Exception:
    pass

# extract-msg (Outlook-.msg-E-Mails) inkl. Abhaengigkeiten mitbuendeln.
for _pkg in ("extract_msg", "compressed_rtf", "RTFDE", "ebcdic"):
    try:
        e_datas, e_binaries, e_hidden = collect_all(_pkg)
        datas += e_datas
        binaries += e_binaries
        hiddenimports += e_hidden
    except Exception:
        pass

a = Analysis(
    [os.path.join("tool", "dokumenten_umbenenner.py")],
    pathex=["tool"],
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
    name="Dokumenten-Umbenenner",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # GUI-Programm: kein Konsolenfenster
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
