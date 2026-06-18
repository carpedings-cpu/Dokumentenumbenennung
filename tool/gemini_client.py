"""
Optionaler Google-Gemini-Modus.

Schickt das Dokument (PDF/Bild als Inline-Daten, sonst Text) an die Gemini-API
und laesst die Felder nach der Verfahrensanweisung bestimmen. Nutzt die
REST-Schnittstelle ueber die Standardbibliothek (kein Zusatzpaket noetig).

Wird nur aufgerufen, wenn der Nutzer den Gemini-Modus aktiviert.
"""

import base64
import json
import os
import urllib.error
import urllib.parse
import urllib.request

MODELL = "gemini-2.5-flash"
_ENDPUNKT = "https://generativelanguage.googleapis.com/v1beta/models/{modell}:generateContent"


def _inline(pfad, mime):
    with open(pfad, "rb") as f:
        return {"inline_data": {"mime_type": mime, "data": base64.standard_b64encode(f.read()).decode("ascii")}}


def _teile(pfad):
    endung = os.path.splitext(pfad)[1].lower()
    if endung == ".pdf":
        return [_inline(pfad, "application/pdf")]
    if endung in (".png", ".jpg", ".jpeg", ".gif", ".webp"):
        mt = {"jpg": "jpeg"}.get(endung[1:], endung[1:])
        return [_inline(pfad, f"image/{mt}")]
    try:
        with open(pfad, "r", encoding="utf-8", errors="replace") as f:
            inhalt = f.read(15000)
    except OSError:
        inhalt = ""
    return [{"text": f"Dateiname: {os.path.basename(pfad)}\n\nInhalt:\n{inhalt}"}]


def analysiere(pfad, va_regeln, api_key, modell=MODELL):
    """Analysiert ein Dokument per Gemini-API und liefert ein Felder-Dict."""
    # Aufgabentext aus dem Claude-Client wiederverwenden (gleiche VA-Logik).
    from api_client import _AUFGABE

    parts = _teile(pfad) + [{"text": _AUFGABE}]
    body = {
        "contents": [{"parts": parts}],
        "system_instruction": {"parts": [{"text": va_regeln}]},
        "generationConfig": {"response_mime_type": "application/json"},
    }
    url = _ENDPUNKT.format(modell=modell) + "?key=" + urllib.parse.quote(api_key)
    req = urllib.request.Request(
        url, data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json"}, method="POST")

    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            daten = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        text = e.read().decode("utf-8", "replace")
        raise RuntimeError(f"Gemini-API-Fehler {e.code}: {text[:300]}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Netzwerkfehler zur Gemini-API: {e}") from e

    try:
        cand = daten["candidates"][0]
        text = "".join(p.get("text", "") for p in cand["content"]["parts"])
    except (KeyError, IndexError) as e:
        raise RuntimeError(f"Unerwartete Gemini-Antwort: {json.dumps(daten)[:300]}") from e

    try:
        return json.loads(text)
    except (ValueError, TypeError) as e:
        raise RuntimeError(f"Gemini-Antwort war kein gültiges JSON:\n{text[:300]}") from e
