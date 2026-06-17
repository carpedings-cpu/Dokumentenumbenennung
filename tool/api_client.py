"""
Optionaler Claude-API-Modus.

Schickt das Dokument (PDF als Dokument-Block, sonst Text) an die Claude-API und
laesst Datum, Quelle, Phase, Dokumententyp, Bezeichnung und Version nach der
Verfahrensanweisung bestimmen. Nutzt das offizielle Anthropic-SDK.

Wird nur aufgerufen, wenn der Nutzer den API-Modus aktiviert. Ohne installiertes
SDK bleibt der Offline-Modus voll funktionsfaehig.
"""

import base64
import json
import os

MODELL = "claude-opus-4-8"

_AUSGABE_SCHEMA = {
    "type": "object",
    "properties": {
        "vorgeschlagener_name": {"type": "string"},
        "datum": {"type": "string"},
        "quelle": {"type": "string"},
        "phase": {"type": "string"},
        "dokumententyp": {"type": "string"},
        "bezeichnung": {"type": "string"},
        "version": {"type": "string"},
        "ist_plan": {"type": "boolean"},
        "hinweis": {"type": "string"},
    },
    "required": [
        "vorgeschlagener_name", "datum", "quelle", "phase", "dokumententyp",
        "bezeichnung", "version", "ist_plan", "hinweis",
    ],
    "additionalProperties": False,
}

_AUFGABE = (
    "Du erhaeltst ein Projektdokument. Bestimme die Felder fuer den Dateinamen "
    "streng nach der oben stehenden Verfahrensanweisung (VA Dokumentenbenennung).\n\n"
    "Regeln fuer die Ausgabe:\n"
    "- datum: Format JJMMTT aus Erstellung/Eingang/Versand. Wenn das Datum laut "
    "VA entfaellt (Plan-, Vertrags-, Betriebsanleitungs-, extern erstellte "
    "Dokumente) oder nicht ermittelbar ist: leer lassen.\n"
    "- quelle: nur bei extern erstellten Dokumenten (Kunde, Planer, Lieferant, "
    "Behoerde), sonst leer.\n"
    "- phase: nur bei Abnahme/Einweisung/IBN, sonst leer.\n"
    "- dokumententyp: MUSS exakt aus der Referenzliste (Kapitel 5) stammen.\n"
    "- bezeichnung: kurze, praezise inhaltliche Beschreibung.\n"
    "- version: nur bei Entwuerfen/freigegebenen Staenden, sonst leer.\n"
    "- ist_plan: true, wenn es eine Planunterlage ist (Kapitel 6).\n"
    "- vorgeschlagener_name: der vollstaendige Dateiname OHNE Endung, exakt nach "
    "Schema gebaut (keine Leerzeichen/Sonderzeichen, Umlaute erlaubt, _ trennt "
    "Felder, - innerhalb eines Feldes). Bei Plaenen das Planbenennungsschema.\n"
    "- hinweis: kurze Begruendung oder Unsicherheit (eine Zeile)."
)


def _sdk_pruefen():
    try:
        import anthropic  # noqa: F401
    except ImportError as e:
        raise RuntimeError(
            "Der API-Modus benoetigt das Anthropic-SDK.\n"
            "Bitte einmalig installieren:  pip install anthropic"
        ) from e


def _dokument_block(pfad):
    endung = os.path.splitext(pfad)[1].lower()
    if endung == ".pdf":
        with open(pfad, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode("ascii")
        return {"type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": b64}}
    if endung in (".png", ".jpg", ".jpeg", ".gif", ".webp"):
        mt = {"jpg": "jpeg"}.get(endung[1:], endung[1:])
        with open(pfad, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode("ascii")
        return {"type": "image",
                "source": {"type": "base64", "media_type": f"image/{mt}", "data": b64}}
    # Textartige Dateien direkt mitgeben
    try:
        with open(pfad, "r", encoding="utf-8", errors="replace") as f:
            inhalt = f.read(15000)
    except OSError:
        inhalt = ""
    return {"type": "text",
            "text": f"Dateiname: {os.path.basename(pfad)}\n\nInhalt:\n{inhalt}"}


def analysiere(pfad, va_regeln, api_key, modell=MODELL):
    """Analysiert ein Dokument per Claude-API und liefert ein Felder-Dict."""
    _sdk_pruefen()
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    block = _dokument_block(pfad)

    resp = client.messages.create(
        model=modell,
        max_tokens=1500,
        system=va_regeln,
        output_config={"format": {"type": "json_schema", "schema": _AUSGABE_SCHEMA}},
        messages=[{"role": "user", "content": [block, {"type": "text", "text": _AUFGABE}]}],
    )

    if getattr(resp, "stop_reason", None) == "refusal":
        raise RuntimeError("Die Anfrage wurde von der API abgelehnt (refusal).")

    text = next((b.text for b in resp.content if getattr(b, "type", None) == "text"), "")
    try:
        return json.loads(text)
    except (ValueError, TypeError) as e:
        raise RuntimeError(f"Antwort der API war kein gueltiges JSON:\n{text[:300]}") from e
