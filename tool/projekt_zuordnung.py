"""
Projekt-Zuordnung fuer den Dokumenten-Umbenenner.

Erkennt aus Dateiname + Inhalt (bzw. Mailtext) das passende Projekt und liefert
den zugehoerigen, real existierenden Projekt-Unterordner zurueck. Damit kann das
Umbenenn-Programm fertige Dateien automatisch in den richtigen Projektordner
einsortieren.

Signale (gleiche Idee wie die Triage):
  - Projektnummer (z. B. 0875 / 0875.20) = starkes, nahezu eindeutiges Signal.
  - Orts-/Stichworte aus dem Ordnernamen und (falls vorhanden) aus
    projekte_mapping.json.

Sicher per Default: einsortiert wird nur in einen Ordner, der bereits unter der
Projektbasis existiert. Bei Unklarheit (kein Treffer / Gleichstand) wird KEIN
Projekt zurueckgegeben -> die Datei bleibt liegen.
"""

import json
import os
import re

# Ordner unter der Projektbasis, die kein Projekt sind.
_IGNORIERT = {"00_posteingang", "triage-berichte", "tool", ".git", "docs",
              "__pycache__", "node_modules"}


def _nummer(text):
    """Fuehrende Projektnummer (3-4 Ziffern) aus einem Namen, sonst ''."""
    m = re.match(r"\s*(\d{3,4})", text or "")
    return m.group(1) if m else ""


def _ziffern(text):
    m = re.match(r"(\d{3,4})", (text or "").strip())
    return m.group(1) if m else ""


def _keywords_aus_name(name):
    """Stichworte aus einem Ordnernamen: Nummer + Worte (>=3 Zeichen)."""
    kws = set()
    num = _nummer(name)
    if num:
        kws.add(num)
    for t in re.split(r"[^0-9A-Za-zÄÖÜäöüß]+", name or ""):
        t = t.strip().lower()
        if len(t) >= 3 and not t.isdigit():
            kws.add(t)
    return kws


def _mapping_kandidaten(base_dir, hier):
    pfade = []
    if base_dir:
        pfade.append(os.path.join(base_dir, "projekte_mapping.json"))
    import sys
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        pfade.append(os.path.join(meipass, "projekte_mapping.json"))
    if hier:
        pfade.append(os.path.join(hier, "projekte_mapping.json"))
        pfade.append(os.path.join(hier, "triage", "projekte_mapping.json"))
    return pfade


def _lade_mapping(base_dir, hier):
    for p in _mapping_kandidaten(base_dir, hier):
        try:
            if os.path.exists(p):
                daten = json.load(open(p, encoding="utf-8"))
                if isinstance(daten, dict) and isinstance(daten.get("projekte"), list):
                    return daten["projekte"]
        except Exception:  # noqa: BLE001
            continue
    return []


def lade_projekte(base_dir, hier=None):
    """Baut die Projektliste aus den realen Unterordnern (+ Mapping-Stichworten)."""
    projekte = []
    by_key = {}
    if base_dir and os.path.isdir(base_dir):
        for name in sorted(os.listdir(base_dir)):
            pfad = os.path.join(base_dir, name)
            if not os.path.isdir(pfad) or name.startswith(".") or name.lower() in _IGNORIERT:
                continue
            nummer = _nummer(name)
            rec = {"ordner": name, "nummer": nummer, "keywords": _keywords_aus_name(name)}
            projekte.append(rec)
            by_key[nummer or name.lower()] = rec

    # Stichworte/Domains aus dem Mapping anreichern – nur fuer existierende Ordner.
    for e in _lade_mapping(base_dir, hier):
        key = _ziffern(e.get("kuerzel", "")) or _nummer(e.get("name", ""))
        rec = by_key.get(key) or by_key.get((e.get("name", "") or "").lower())
        if not rec:
            continue
        for kw in e.get("betreff_stichworte", []):
            kw = (kw or "").strip().lower()
            if (len(kw) >= 3 and not kw.isdigit()) or kw == rec["nummer"]:
                rec["keywords"].add(kw)
        for d in e.get("domains", []):
            d = (d or "").strip().lower()
            if len(d) >= 3:
                rec["keywords"].add(d)
    return projekte


def finde_projekt(projekte, text):
    """Bestes Projekt fuer den Text – oder None, wenn unklar/mehrdeutig."""
    # Unterstriche zu Leerzeichen: in Dateinamen wie 260617_0875_... ist die
    # Projektnummer von '_' umschlossen; '_' zaehlt sonst als Wortzeichen und
    # verhindert den \b-Treffer.
    low = re.sub(r"_+", " ", (text or "").lower())
    if not low.strip() or not projekte:
        return None
    best = None
    best_s = second_s = 0
    best_num = False
    for p in projekte:
        s = 0
        nummer_hit = False
        if p["nummer"] and re.search(r"\b" + re.escape(p["nummer"]) + r"\b", low):
            s += 5
            nummer_hit = True
        for kw in p["keywords"]:
            if kw == p["nummer"]:
                continue
            if re.search(r"\b" + re.escape(kw) + r"\b", low):
                s += 1
        if s > best_s:
            second_s, best_s, best, best_num = best_s, s, p, nummer_hit
        elif s > second_s:
            second_s = s
    if best is None or best_s == 0:
        return None
    if best_num:
        # Genau eine Projektnummer im Text -> sicher; zwei Nummern -> mehrdeutig.
        return best if second_s < 5 else None
    if best_s >= 2 and best_s > second_s:
        return best
    return None


def projektordner(base_dir, name):
    """Vollpfad eines Projektordners, falls er existiert – sonst ''."""
    if not base_dir or not name:
        return ""
    pfad = os.path.join(base_dir, name)
    return pfad if os.path.isdir(pfad) else ""
