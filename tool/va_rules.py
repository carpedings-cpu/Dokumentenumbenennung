"""
Regeln der KPC-Verfahrensanweisung "Dokumentenbenennung & Versionierung"
(Version 1.3, Stand Juni 2026).

Maschinen-Referenz fuer das Umbenennungs-Tool. Inhaltlich identisch mit
.claude/skills/dokumentenumbenennung/SKILL.md - beide muessen synchron bleiben.
"""

import re
import unicodedata

# ---------------------------------------------------------------------------
# Verbindliche Dokumententypen (Referenzliste Kapitel 5, VA 1.3)
# Platzhalter wie _[Lieferant]/_[Bezeichnung] sind entfernt; den variablen Teil
# traegt der Nutzer in der Bezeichnung nach. Feste Unterstrich-Teile (z. B.
# Vertragsbedingungen_BVB, AG_Vorgabe_Montage) bleiben Teil des Typs.
# ---------------------------------------------------------------------------
DOKUMENTTYPEN = [
    # 5.1 Allgemeine Dokumententypen
    "Vermerk", "Gesprächsnotiz", "Protokoll", "Schriftverkehr", "E-Mail",
    "Bericht", "Anweisung", "Entscheidung", "Versandnachweis", "Datenblatt",
    "Präsentation",
    # 5.2 Organisation
    "Verfahrensanweisung", "Arbeitsanweisung", "Anleitung", "Schulungsunterlagen",
    # 5.3 Vertragsmanagement – Vertragliche Grundlagen
    "Ausschreibung_Leistungsverzeichnis", "Ausschreibung_Bieterfragen",
    "Ausschreibung_Kalkulation", "Ausschreibung_Angebot_signiert",
    "Ausschreibung_Alternativangebot_signiert", "Vergabeprotokoll",
    "Vergabe_Auftragsschreiben", "AG_BZP", "AG_Logistikhandbuch",
    "Vertrag_Auftraggeber", "Vertragsbedingungen_BVB", "Vertragsbedingungen_ZVB",
    "Vertragsbedingungen_TVB", "Vertragsbedingungen_AVB", "Stückliste",
    # 5.4 Vertragsmanagement – Technische Unterlagen (vom Auftraggeber)
    "AG_Planunterlagen", "AG_Technische-Stückliste", "AG_Herstellerunterlagen",
    "AG_Produktdatenblätter", "AG_Bemusterungsunterlagen",
    "AG_Vorschriften_Baugenehmigung", "AG_Vorschriften_Brandschutznachweis",
    "AG_Nachweise_Zertifizierung", "AG_Vorgabe_Montage", "AG_Vorgabe_Revision",
    "AG_Vorgabe_Wartung", "AG_Vorgabe_Sicherheit", "AG_Vorgabe_Logistik",
    "Muster",
    # 5.5 Vertragsmanagement – Projektmanagement
    "Richtlinie", "Baustellenordnung", "Nachweisdokument", "Prüfbericht",
    "Formblatt", "Behinderungsanzeige", "Vertrag", "Vorgabe",
    # 5.6 Wartungsvertrag
    "Wartungsvertrag",
    # 5.8 Inbetriebnahme
    "IBN_Anmeldung", "IBN_Protokoll",
    # 5.9 Abnahme
    "Abnahme_Leistungsfeststellung", "Abnahme_Anmeldung",
    "Abnahme_Terminbestätigung", "Abnahme_Stückliste", "Abnahmeprotokoll",
    "Abnahme_Prüfprotokoll", "Abnahme_Fertigstellungsanzeige",
    "Abnahme_Mängelliste", "Abnahme_Mängelfreimeldung",
    # 5.10 Betrieb, Service, Wartung
    "Revisionsunterlagen", "Betriebsanleitung", "Wartungsvorgabe",
    "Wartungsprotokoll", "Servicebericht", "Serviceanforderung", "Mangelanzeige",
    # 5.11 Einkauf | Werksplan und Ueberwachung der Lieferung
    "Auftragsbestätigung", "Lieferavis", "Lieferterminänderung", "Werksplan",
]

# Typen, bei denen das Datumsfeld laut VA entfaellt (Vertrags-/AG-Unterlagen,
# Betriebsanleitungen/Wartungsvorgaben, extern erstellte Einkaufsdokumente).
TYPEN_OHNE_DATUM = {
    # Vertragliche Grundlagen (5.3)
    "Ausschreibung_Leistungsverzeichnis", "Ausschreibung_Bieterfragen",
    "Ausschreibung_Kalkulation", "Ausschreibung_Angebot_signiert",
    "Ausschreibung_Alternativangebot_signiert", "Vergabeprotokoll",
    "Vergabe_Auftragsschreiben", "AG_BZP", "AG_Logistikhandbuch",
    "Vertrag_Auftraggeber", "Vertragsbedingungen_BVB", "Vertragsbedingungen_ZVB",
    "Vertragsbedingungen_TVB", "Vertragsbedingungen_AVB", "Stückliste",
    # Technische Unterlagen vom AG (5.4)
    "AG_Planunterlagen", "AG_Technische-Stückliste", "AG_Herstellerunterlagen",
    "AG_Produktdatenblätter", "AG_Bemusterungsunterlagen",
    "AG_Vorschriften_Baugenehmigung", "AG_Vorschriften_Brandschutznachweis",
    "AG_Nachweise_Zertifizierung", "AG_Vorgabe_Montage", "AG_Vorgabe_Revision",
    "AG_Vorgabe_Wartung", "AG_Vorgabe_Sicherheit", "AG_Vorgabe_Logistik",
    "Muster",
    # Projektmanagement / Verträge (5.5, 5.6)
    "Richtlinie", "Baustellenordnung", "Nachweisdokument", "Prüfbericht",
    "Formblatt", "Behinderungsanzeige", "Vertrag", "Vorgabe", "Wartungsvertrag",
    # Betrieb (5.10) – Gerät/Anlage als Bezeichnung
    "Betriebsanleitung", "Wartungsvorgabe",
    # Einkauf (5.11) – Quelle statt Datum
    "Auftragsbestätigung", "Lieferavis", "Lieferterminänderung", "Werksplan",
}

# Extern erstellte Typen (Kap. 5.11) -> Quelle gehoert in den Namen.
TYPEN_MIT_QUELLE = {
    "Auftragsbestätigung", "Lieferavis", "Lieferterminänderung", "Werksplan",
}

# ---------------------------------------------------------------------------
# Planbenennungsschema (Kapitel 6) - Kodierungen
# ---------------------------------------------------------------------------
PLAN_AUFTRAGSART = {"10": "Planung", "20": "Ausführung"}
PLAN_LEISTUNGSPHASE = {
    "1": "Grundlagenermittlung", "2": "Vorplanung", "3": "Entwurfsplanung",
    "4": "Genehmigungsplanung", "5": "Ausführungsplanung",
    "6": "Vorbereitung der Vergabe", "7": "Mitwirkung bei der Vergabe",
    "8": "Objektüberwachung/Ausführung", "9": "Objektbetreuung",
}
PLAN_PLANINHALT = {
    "A": "Ansichten", "B": "Bodenplan", "D": "Details", "E": "Einrichtung",
    "F": "Flächenkonzept", "G": "Gründung/Fundament", "I": "Installation",
    "K": "Kälte", "L": "Lüftungsdecken/Hauben", "Ü": "Übersichtsplan",
    "W": "Wände",
}
PLAN_GESCHOSS = {
    "U2": "2. UG", "U1": "1. UG", "EG": "Erdgeschoss", "01": "1. OG",
    "02": "2. OG", "03": "3. OG", "DG": "Dachgeschoss", "99": "geschossübergreifend",
}
PLAN_STATUS = {"V": "Vorabzug", "PL": "Prüflauf", "F": "Freigegeben"}

# ---------------------------------------------------------------------------
# Datum
# ---------------------------------------------------------------------------
_MONATE = {
    # Deutsch (voll + Abkürzungen)
    "januar": 1, "jan": 1, "februar": 2, "feb": 2, "märz": 3, "maerz": 3,
    "mrz": 3, "april": 4, "apr": 4, "mai": 5, "juni": 6, "jun": 6,
    "juli": 7, "jul": 7, "august": 8, "aug": 8, "september": 9, "sep": 9,
    "sept": 9, "oktober": 10, "okt": 10, "november": 11, "nov": 11,
    "dezember": 12, "dez": 12,
    # Englisch (für E-Mail-Datumszeilen wie "15 Apr 2025")
    "january": 1, "february": 2, "march": 3, "may": 5, "june": 6, "july": 7,
    "october": 10, "december": 12, "mar": 3, "oct": 10, "dec": 12,
}


def zu_jjmmtt(tag, monat, jahr):
    """Baut JJMMTT aus Tag/Monat/Jahr (Jahr 2- oder 4-stellig)."""
    jahr = int(jahr)
    if jahr > 99:
        jahr = jahr % 100
    return f"{jahr:02d}{int(monat):02d}{int(tag):02d}"


def datum_aus_text(text):
    """Sucht das erste plausible Datum im Text und gibt es als JJMMTT zurueck."""
    if not text:
        return ""
    # 2025-04-15 / 25-04-15
    m = re.search(r"\b(\d{2,4})-(\d{1,2})-(\d{1,2})\b", text)
    if m:
        j, mo, t = m.group(1), m.group(2), m.group(3)
        if 1 <= int(mo) <= 12 and 1 <= int(t) <= 31:
            return zu_jjmmtt(t, mo, j)
    # 15.04.2025 / 15.04.25 / 15.4.2025
    m = re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{2,4})\b", text)
    if m:
        t, mo, j = m.group(1), m.group(2), m.group(3)
        if 1 <= int(mo) <= 12 and 1 <= int(t) <= 31:
            return zu_jjmmtt(t, mo, j)
    # 15. April 2025
    m = re.search(r"\b(\d{1,2})\.?\s+([A-Za-zäöüÄÖÜ]+)\s+(\d{4})\b", text)
    if m and m.group(2).lower() in _MONATE:
        return zu_jjmmtt(m.group(1), _MONATE[m.group(2).lower()], m.group(3))
    return ""


# ---------------------------------------------------------------------------
# Typ-Erkennung (regelbasiert / offline)
# ---------------------------------------------------------------------------
def typ_aus_text(text):
    """Schlaegt den am besten passenden Dokumententyp aus der Referenzliste vor."""
    if not text:
        return ""
    # Bindestriche/Unterstriche zu Leerzeichen -> ganzwoertliche Treffer moeglich
    low = re.sub(r"[-_]+", " ", text.lower())
    treffer = []
    for typ in DOKUMENTTYPEN:
        begriff = typ.replace("_", " ").replace("-", " ").lower()
        if re.search(r"\b" + re.escape(begriff) + r"\b", low):
            treffer.append((len(begriff), typ))
    if treffer:
        treffer.sort(reverse=True)   # laengster (spezifischster) Treffer gewinnt
        return treffer[0][1]
    return ""


# Synonyme/haeufige Begriffe -> erlaubter VA-Typ. Bewusst konservativ gehalten.
_SYNONYME = {
    "angebot": "Schriftverkehr", "kostenvoranschlag": "Schriftverkehr",
    "brief": "Schriftverkehr", "schreiben": "Schriftverkehr",
    "anschreiben": "Schriftverkehr", "fax": "Schriftverkehr",
    "korrespondenz": "Schriftverkehr",
    "email": "E-Mail", "e mail": "E-Mail", "mail": "E-Mail",
    "aktennotiz": "Vermerk", "notiz": "Vermerk",
    "lieferschein": "Lieferavis", "auftragsbestaetigung": "Auftragsbestätigung",
    # VA 1.3: neue/umbenannte Typen
    "praesentation": "Präsentation", "powerpoint": "Präsentation",
    "ppt": "Präsentation", "folien": "Präsentation",
    "wartungsanforderung": "Serviceanforderung",   # 1.1 -> 1.3 umbenannt
    "maengelanzeige": "Mangelanzeige", "mangelmeldung": "Mangelanzeige",
    "wartungsbericht": "Wartungsprotokoll",
    "leistungsverzeichnis": "Ausschreibung_Leistungsverzeichnis",
    "kalkulation": "Ausschreibung_Kalkulation",
    "vergabe": "Vergabeprotokoll",
    "arbeitsanweisung": "Arbeitsanweisung", "va": "Verfahrensanweisung",
}


def normalisiere_typ(typ):
    """Bildet einen Typ auf die VA-Referenzliste ab; '' wenn nicht zulaessig."""
    if not typ:
        return ""
    roh = typ.strip()
    for t in DOKUMENTTYPEN:                       # exakter Treffer (case-insensitiv)
        if t.lower() == roh.lower():
            return t
    schluessel = re.sub(r"[-_]+", " ", roh.lower()).replace("ä", "ae")\
        .replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return _SYNONYME.get(schluessel, _SYNONYME.get(roh.lower(), ""))


def normalisiere_datum(s):
    """Bringt eine Datumsangabe in das Format JJMMTT (sonst '')."""
    s = (s or "").strip()
    if re.fullmatch(r"\d{6}", s):                 # bereits JJMMTT
        return s
    if re.fullmatch(r"\d{8}", s):                 # JJJJMMTT
        return s[2:]
    return datum_aus_text(s)


# ---------------------------------------------------------------------------
# Bereinigung & Namensbau
# ---------------------------------------------------------------------------
_ERLAUBT = re.compile(r"[^0-9A-Za-zÄÖÜäöüß_.\-]")


def bereinige(feld):
    """Entfernt Leerzeichen/Sonderzeichen; Umlaute bleiben erhalten."""
    if not feld:
        return ""
    feld = unicodedata.normalize("NFC", str(feld)).strip()
    feld = re.sub(r"\s+", "-", feld)        # Leerzeichen -> Bindestrich
    feld = _ERLAUBT.sub("", feld)            # uebrige Sonderzeichen weg
    feld = re.sub(r"-{2,}", "-", feld)       # mehrfache Bindestriche
    return feld.strip("-_")


def baue_standardname(felder):
    """JJMMTT_[Quelle]_[Phase]_Dokumententyp_Bezeichnung[_Version] (ohne Endung)."""
    teile = []
    for key in ("datum", "quelle", "phase", "dokumententyp", "bezeichnung"):
        wert = bereinige(felder.get(key, ""))
        if wert:
            teile.append(wert)
    name = "_".join(teile)
    version = bereinige(felder.get("version", ""))
    if version:
        name = f"{name}_{version}" if name else version
    return name


def baue_planname(felder):
    """[Projektnr].[Auftragsart]_[LP]_[Inhalt]_[Geschoss]_[Plannr]_[Index]_[Status]."""
    proj = bereinige(felder.get("projektnr", ""))
    art = bereinige(felder.get("auftragsart", ""))
    kopf = f"{proj}.{art}" if proj and art else (proj or art)
    rest = [bereinige(felder.get(k, "")) for k in
            ("leistungsphase", "planinhalt", "geschoss", "plannr", "index", "status")]
    rest = [r for r in rest if r]
    teile = [kopf] + rest if kopf else rest
    teile = [t for t in teile if t]
    return "_".join(teile)


def eindeutiger_zielname(ordner, neuer_stamm, endung, original_name, vorhandene=None):
    """
    Liefert einen kollisionsfreien Zielnamen. Ueberschreibt nie.
    `vorhandene` = optionales Set bereits vergebener Namen (fuer Stapelbetrieb).
    """
    import os
    if vorhandene is None:
        vorhandene = set()
    kandidat = f"{neuer_stamm}{endung}"
    if kandidat == original_name:
        return kandidat  # keine Aenderung
    i = 2
    while (os.path.exists(os.path.join(ordner, kandidat)) and kandidat != original_name) \
            or kandidat.lower() in vorhandene:
        kandidat = f"{neuer_stamm}-{i:02d}{endung}"
        i += 1
    return kandidat
