#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPC Posteingangs-Triage für Outlook (Windows, COM/pywin32).

Aufgabe (dockt an den Dokumentenbenennungs-Skill an, baut nichts doppelt):
  Outlook-Posteingang LESEN -> klassifizieren -> relevante Mails als .msg in den
  Eingangsordner des Dokumentenbenennungs-Skills legen + HTML-Übersicht erzeugen.
  Die Ablage/Benennung (Mailtext-PDF + Anhänge, Projekt-Unterordner) macht
  weiterhin der Skill.

SICHERHEIT (nicht verhandelbar):
  - Posteingang ist READ-ONLY. Es wird NICHTS verschoben, gelöscht oder als
    gelesen markiert. Eigenschaften lesen und .msg exportieren ändern den
    Posteingang nicht. Optional kann (nur mit --mit-kategorie) eine reversible
    Outlook-Kategorie gesetzt werden – das ist die einzige erlaubte Schreibaktion.
  - Inkrementell: nur Mails seit dem letzten scharfen Lauf (Marker in
    triage_state.json: letzte ReceivedTime + verarbeitete EntryIDs).

DSGVO / Klassifizierung zweistufig:
  - Stufe 1: lokal & regelbasiert (Absender-Domain + Betreff-Stichworte gegen
    projekte_mapping.json). Lokal eindeutige Mails gehen NIE an eine API.
  - Stufe 2 (optional, abschaltbar): nur lokal NICHT eindeutige Mails gehen an
    die Claude-API. Schlüssel aus .env, standardmäßig AUS.

ABLAUF:
  - Standard = TROCKENLAUF: nur klassifizieren + Übersicht bauen, nichts in den
    Eingangsordner schreiben, Marker NICHT weiterstellen.
  - Erst mit --scharf werden .msg abgelegt und der Marker weitergestellt.

Aufruf:
  python triage.py                # Trockenlauf (empfohlen zuerst)
  python triage.py --scharf       # Ablage scharf schalten
  python triage.py --scharf --mit-kategorie
"""

import argparse
import datetime as dt
import html
import json
import os
import re
import sys

HIER = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# Konfiguration (per triage_config.json überschreibbar)
# ---------------------------------------------------------------------------
CONFIG = {
    "base_dir": r"C:\Users\ziegler\Desktop\Dokumentenumbenennung",
    "eingang_unterordner": "00_Posteingang",     # Übergabe an den Skill
    "bericht_unterordner": "Triage-Berichte",
    "outlook_kategorie": "KPC-Triage",            # nur mit --mit-kategorie
    # Ordner, die NICHT als Projekt zählen:
    "ignorierte_ordner": ["00_Posteingang", "Triage-Berichte", "tool",
                          ".git", "docs", "__pycache__"],
}


def lade_config():
    pfad = os.path.join(HIER, "triage_config.json")
    if os.path.exists(pfad):
        try:
            CONFIG.update(json.load(open(pfad, encoding="utf-8")))
        except Exception as e:  # noqa: BLE001
            print(f"WARN: triage_config.json nicht lesbar: {e}")
    return CONFIG


def lade_env(pfad):
    werte = {}
    if os.path.exists(pfad):
        for zeile in open(pfad, encoding="utf-8"):
            zeile = zeile.strip()
            if not zeile or zeile.startswith("#") or "=" not in zeile:
                continue
            k, v = zeile.split("=", 1)
            werte[k.strip()] = v.strip().strip('"').strip("'")
    return werte


# ---------------------------------------------------------------------------
# Kategorien & Dringlichkeit
# ---------------------------------------------------------------------------
KATEGORIEN = [
    "Auftragsbestaetigung-mit-Terminaenderung",
    "Lieferavis",
    "Maengel/Behinderung",
    "Nachtrag",
    "Rechnung",
    "Info",
]

# Stichworte je Kategorie (lower-case Teilstrings im Betreff/Auszug)
KAT_STICHWORTE = {
    "Maengel/Behinderung": ["mangel", "mängel", "maengel", "behinderung",
                            "bedenken", "bedenkenanmeldung", "verzug",
                            "behinderungsanzeige"],
    "Nachtrag": ["nachtrag", "nachtragsangebot", "na-angebot"],
    "Lieferavis": ["lieferavis", "avis", "liefermitteilung", "versandanzeige",
                   "lieferankündigung", "lieferankuendigung", "versandavis"],
    "Rechnung": ["rechnung", "invoice", "gutschrift", "zahlungsaufforderung",
                 "mahnung"],
}
# Auftragsbestätigung-mit-Terminänderung wird gesondert erkannt (AB + Termin).
AB_WORTE = ["auftragsbestätigung", "auftragsbestaetigung", "auftragsbest"]
TERMIN_WORTE = ["termin", "liefertermin", "verschoben", "verzug", "verzögerung",
                "verzoegerung", "neuer termin", "terminverschiebung",
                "terminänderung", "terminaenderung"]

DRINGLICHKEIT = {
    "Maengel/Behinderung": "Hoch",
    "Auftragsbestaetigung-mit-Terminaenderung": "Hoch",
    "Nachtrag": "Hoch",
    "Rechnung": "Mittel",
    "Lieferavis": "Mittel",
    "Info": "Niedrig",
}
GRUPPEN_REIHENFOLGE = ["Hoch", "Mittel", "Niedrig"]


# ---------------------------------------------------------------------------
# Projektliste / Mapping
# ---------------------------------------------------------------------------
def mapping_pfad():
    return os.path.join(HIER, "projekte_mapping.json")


def lade_mapping():
    pfad = mapping_pfad()
    if os.path.exists(pfad):
        try:
            daten = json.load(open(pfad, encoding="utf-8"))
            return daten if isinstance(daten, dict) else {"projekte": []}
        except Exception as e:  # noqa: BLE001
            print(f"WARN: projekte_mapping.json nicht lesbar: {e}")
    return {"projekte": []}


def speichere_mapping(mapping):
    json.dump(mapping, open(mapping_pfad(), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


def _kuerzel_aus_ordner(name):
    m = re.match(r"\s*(\d{3,})", name)         # führende Projektnummer
    if m:
        return m.group(1)
    wort = re.sub(r"[^0-9A-Za-zÄÖÜäöüß]+", "", name)[:6]
    return wort.upper() or "PROJ"


def aktualisiere_projekte_aus_ordnern(mapping, cfg):
    """Ergänzt das Mapping um neue Projekt-Unterordner unter base_dir."""
    base = cfg["base_dir"]
    if not os.path.isdir(base):
        print(f"HINWEIS: base_dir nicht gefunden: {base} "
              "(Projektliste wird nicht automatisch ergänzt).")
        return mapping
    bekannt = {(p.get("ordner") or p.get("name") or "").lower()
               for p in mapping["projekte"]}
    ignor = {x.lower() for x in cfg.get("ignorierte_ordner", [])}
    neu = 0
    for name in sorted(os.listdir(base)):
        pfad = os.path.join(base, name)
        if not os.path.isdir(pfad) or name.startswith(".") or name.lower() in ignor:
            continue
        if name.lower() in bekannt:
            continue
        mapping["projekte"].append({
            "name": name,
            "kuerzel": _kuerzel_aus_ordner(name),
            "ordner": name,
            "domains": [],
            "betreff_stichworte": [t for t in re.split(r"[^0-9A-Za-zÄÖÜäöüß]+", name) if len(t) >= 3],
        })
        neu += 1
    if neu:
        speichere_mapping(mapping)
        print(f"projekte_mapping.json: {neu} neue Projekt(e) aus Ordnern ergänzt "
              "(Domains/Stichworte bitte selbst pflegen).")
    return mapping


def finde_projekt(mapping, domain, betreff):
    """Stufe-1-Projektzuordnung über Domain bzw. Betreff-Stichworte."""
    betreff_low = (betreff or "").lower()
    domain = (domain or "").lower()
    bestes = None
    bestscore = 0
    for p in mapping["projekte"]:
        score = 0
        for d in p.get("domains", []):
            d = (d or "").lower().strip()
            if d and (domain == d or domain.endswith("." + d)):
                score += 3
        for kw in p.get("betreff_stichworte", []):
            kw = (kw or "").lower().strip()
            if kw and re.search(r"\b" + re.escape(kw) + r"\b", betreff_low):
                score += 1
        if score > bestscore:
            bestscore, bestes = score, p
    return (bestes, bestscore) if bestes else (None, 0)


def finde_kategorie(betreff, auszug):
    text = f"{betreff or ''} {auszug or ''}".lower()
    if any(w in text for w in AB_WORTE) and any(w in text for w in TERMIN_WORTE):
        return "Auftragsbestaetigung-mit-Terminaenderung"
    for kat in ("Maengel/Behinderung", "Nachtrag", "Rechnung", "Lieferavis"):
        if any(w in text for w in KAT_STICHWORTE[kat]):
            return kat
    return "Info"


# ---------------------------------------------------------------------------
# Zustand (Marker für inkrementelle Verarbeitung)
# ---------------------------------------------------------------------------
def state_pfad():
    return os.path.join(HIER, "triage_state.json")


def lade_state():
    pfad = state_pfad()
    if os.path.exists(pfad):
        try:
            return json.load(open(pfad, encoding="utf-8"))
        except Exception:  # noqa: BLE001
            pass
    return {"letzte_received": None, "verarbeitete_entry_ids": []}


def speichere_state(state):
    state["verarbeitete_entry_ids"] = state.get("verarbeitete_entry_ids", [])[-4000:]
    json.dump(state, open(state_pfad(), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Outlook (READ-ONLY)
# ---------------------------------------------------------------------------
def outlook_posteingang():
    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError("pywin32 wird benötigt: pip install pywin32") from e
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return ns.GetDefaultFolder(6)  # 6 = olFolderInbox


def smtp_adresse(item):
    """Liefert die SMTP-Absenderadresse (auch bei Exchange) – nur lesend."""
    try:
        if (item.SenderEmailType or "").upper() == "EX":
            try:
                return item.Sender.GetExchangeUser().PrimarySmtpAddress or ""
            except Exception:
                PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                try:
                    return item.PropertyAccessor.GetProperty(PR_SMTP) or ""
                except Exception:
                    return item.SenderEmailAddress or ""
        return item.SenderEmailAddress or ""
    except Exception:
        return ""


def py_datetime(com_zeit):
    """pywin32 COM-Zeit -> naive datetime."""
    try:
        return dt.datetime(com_zeit.year, com_zeit.month, com_zeit.day,
                           com_zeit.hour, com_zeit.minute, com_zeit.second)
    except Exception:
        return dt.datetime.fromtimestamp(0)


def sicherer_name(text, maxlen=80):
    text = re.sub(r"[\\/:*?\"<>|\r\n\t]+", "_", (text or "").strip())
    text = re.sub(r"\s+", "-", text)
    return (text[:maxlen] or "Mail").strip("-_")


# ---------------------------------------------------------------------------
# Stufe 2: Claude-API (optional, nur für unklare Mails)
# ---------------------------------------------------------------------------
def klassifiziere_stufe2(meta, api_key, projektnamen):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    schema = {
        "type": "object",
        "properties": {
            "projekt": {"type": "string"},
            "kategorie": {"type": "string", "enum": KATEGORIEN},
        },
        "required": ["projekt", "kategorie"],
        "additionalProperties": False,
    }
    system = (
        "Du ordnest eine eingehende Projekt-E-Mail einem Projekt und einer "
        "Kategorie zu. Wähle die Kategorie GENAU aus der erlaubten Liste. "
        "Bekannte Projekte (Kürzel – Name):\n"
        + "\n".join(projektnamen) +
        "\nWenn kein Projekt passt, gib projekt = '' zurück."
    )
    user = (f"Absender: {meta['absender']}\nBetreff: {meta['betreff']}\n\n"
            f"Auszug:\n{meta['auszug']}")
    resp = client.messages.create(
        model="claude-opus-4-8", max_tokens=400, system=system,
        output_config={"format": {"type": "json_schema", "schema": schema}},
        messages=[{"role": "user", "content": user}])
    if getattr(resp, "stop_reason", None) == "refusal":
        raise RuntimeError("API-Klassifizierung abgelehnt (refusal).")
    text = next((b.text for b in resp.content if getattr(b, "type", None) == "text"), "")
    daten = json.loads(text)
    return daten.get("projekt", ""), daten.get("kategorie", "Info")


# ---------------------------------------------------------------------------
# HTML-Übersicht (KPC-Design, A4 quer)
# ---------------------------------------------------------------------------
def baue_html(zeilen, scharf):
    def esc(x):
        return html.escape(str(x or ""))

    nach_gruppe = {g: [] for g in GRUPPEN_REIHENFOLGE}
    for z in zeilen:
        nach_gruppe.get(z["gruppe"], nach_gruppe["Niedrig"]).append(z)

    stand = dt.datetime.now().strftime("%d.%m.%Y %H:%M")
    modus = "Scharf (Ablage aktiv)" if scharf else "Trockenlauf (keine Ablage)"

    teile = ["""<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">
<title>KPC Posteingangs-Triage</title>
<link href="https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;700&display=swap" rel="stylesheet">
<style>
  @page { size: A4 landscape; margin: 12mm; }
  * { box-sizing: border-box; }
  body { font-family:'Roboto Condensed','Arial Narrow',sans-serif; font-weight:300;
         color:#2d2926; margin:0; padding:24px; background:#ffffff; }
  header { border-bottom:3px solid #c8a882; padding-bottom:10px; margin-bottom:18px; }
  h1 { font-weight:300; font-size:24px; letter-spacing:1px; margin:0; color:#2d2926; }
  .meta { color:#6b5040; font-size:12px; margin-top:4px; }
  h2 { font-weight:400; font-size:15px; color:#ffffff; background:#2d2926;
       padding:6px 10px; margin:18px 0 0 0; border-left:6px solid #c8a882; }
  h2.mittel { background:#6b5040; }
  h2.niedrig { background:#9a8a78; }
  h2.unklar { background:#8a2f2f; }
  table { width:100%; border-collapse:collapse; margin-top:0; font-size:12px; }
  th { text-align:left; background:#6b5040; color:#fff; font-weight:400;
       padding:6px 8px; }
  td { padding:5px 8px; border-bottom:1px solid #e6ddcf; vertical-align:top; }
  tr:nth-child(even) td { background:#faf6ef; }
  .kuerzel { color:#6b5040; font-weight:700; }
  .ki { font-size:10px; color:#8a2f2f; }
  footer { margin-top:20px; color:#6b5040; font-size:11px;
           border-top:1px solid #c8a882; padding-top:8px; }
</style></head><body>
<header>
  <h1>KPC &middot; Posteingangs-Triage</h1>
  <div class="meta">Stand: """ + esc(stand) + " &nbsp;|&nbsp; Modus: " + esc(modus)
        + " &nbsp;|&nbsp; Mails: " + str(len(zeilen)) + """</div>
</header>"""]

    klasse = {"Hoch": "", "Mittel": "mittel", "Niedrig": "niedrig"}
    titel = {"Hoch": "Dringlichkeit: Hoch",
             "Mittel": "Dringlichkeit: Mittel",
             "Niedrig": "Info / Sonstiges"}
    for g in GRUPPEN_REIHENFOLGE:
        rows = nach_gruppe[g]
        if not rows:
            continue
        teile.append(f'<h2 class="{klasse[g]}">{esc(titel[g])} ({len(rows)})</h2>')
        teile.append("<table><tr><th>Projekt</th><th>Kategorie</th>"
                     "<th>Absender</th><th>Betreff</th><th>Eingang</th></tr>")
        for z in rows:
            ki = ' <span class="ki">(KI)</span>' if z.get("ki") else ""
            teile.append(
                "<tr>"
                f'<td><span class="kuerzel">{esc(z["kuerzel"])}</span> {esc(z["projekt"])}</td>'
                f"<td>{esc(z['kategorie'])}{ki}</td>"
                f"<td>{esc(z['absender'])}</td>"
                f"<td>{esc(z['betreff'])}</td>"
                f"<td>{esc(z['eingang'])}</td>"
                "</tr>")
        teile.append("</table>")

    teile.append('<footer>READ-ONLY-Triage &middot; im scharfen Lauf werden '
                 '<b>alle</b> Mails als .msg an den Dokumentenbenennungs-Skill '
                 'übergeben und dort (Mailtext-PDF + Anhänge) umbenannt. '
                 'Die Gruppierung dient nur der Übersicht.</footer></body></html>')
    return "".join(teile)


# ---------------------------------------------------------------------------
# Hauptlauf
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="KPC Posteingangs-Triage (read-only).")
    ap.add_argument("--scharf", action="store_true",
                    help="Ablage scharf schalten (.msg exportieren, Marker weiterstellen).")
    ap.add_argument("--mit-kategorie", action="store_true",
                    help="Reversible Outlook-Kategorie setzen (nur mit --scharf).")
    ap.add_argument("--stufe2", action="store_true",
                    help="Stufe-2-API erzwingen (sonst aus .env: TRIAGE_USE_API).")
    ap.add_argument("--seit", default=None,
                    help="Startdatum YYYY-MM-DD (überschreibt den Marker einmalig).")
    ap.add_argument("--heute", action="store_true",
                    help="Nur die heutigen Mails ansehen (Marker ignorieren, zum erneuten Prüfen).")
    ap.add_argument("--max", type=int, default=0, help="Maximale Anzahl Mails (Debug).")
    args = ap.parse_args()

    cfg = lade_config()
    env = lade_env(os.path.join(HIER, ".env"))
    use_api = args.stufe2 or str(env.get("TRIAGE_USE_API", "")).lower() in ("1", "true", "ja", "yes")
    api_key = env.get("ANTHROPIC_API_KEY", "").strip()
    if use_api and not api_key:
        print("HINWEIS: Stufe 2 aktiv, aber kein ANTHROPIC_API_KEY in .env – Stufe 2 wird übersprungen.")
        use_api = False

    base = cfg["base_dir"]
    eingang = os.path.join(base, cfg["eingang_unterordner"])
    bericht_dir = os.path.join(base, cfg["bericht_unterordner"])

    mapping = aktualisiere_projekte_aus_ordnern(lade_mapping(), cfg)
    projektnamen = [f"{p.get('kuerzel','')} – {p.get('name','')}" for p in mapping["projekte"]]

    heute_anfang = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    state = lade_state()
    if args.heute:
        marker = heute_anfang
        print("--heute: betrachte alle heutigen Mails erneut (Marker wird nicht genutzt).")
    elif args.seit:
        marker = dt.datetime.strptime(args.seit, "%Y-%m-%d")
    elif state.get("letzte_received"):
        marker = dt.datetime.fromisoformat(state["letzte_received"])
    else:
        # Erster Lauf: erst AB HEUTE beginnen (keine alten Mails einsammeln).
        marker = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        print(f"Erster Lauf: betrachte nur Mails ab heute ({marker:%d.%m.%Y}).")
    verarbeitet = set(state.get("verarbeitete_entry_ids", []))

    print(f"Modus: {'SCHARF' if args.scharf else 'TROCKENLAUF'} | "
          f"Stufe 2 (API): {'AN' if use_api else 'AUS'} | ab {marker:%d.%m.%Y %H:%M}")

    posteingang = outlook_posteingang()
    items = posteingang.Items
    items.Sort("[ReceivedTime]", True)   # absteigend (neueste zuerst)

    zeilen = []
    neue_max_received = marker
    n_export = 0
    geprueft = 0

    item = items.GetFirst()
    while item is not None:
        if args.max and geprueft >= args.max:
            break
        try:
            if int(getattr(item, "Class", 0)) != 43:   # 43 = olMail
                item = items.GetNext(); continue
            received = py_datetime(item.ReceivedTime)
            if received <= marker:
                break                                    # ab hier nur Älteres
            entry_id = item.EntryID
            if entry_id in verarbeitet:
                item = items.GetNext(); continue
            geprueft += 1

            betreff = item.Subject or ""
            absender_smtp = smtp_adresse(item)
            domain = absender_smtp.split("@")[-1] if "@" in absender_smtp else ""
            try:
                auszug = (item.Body or "")[:800]
            except Exception:
                auszug = ""

            # --- Stufe 1: lokal/regelbasiert ---
            projekt, score = finde_projekt(mapping, domain, betreff)
            kategorie = finde_kategorie(betreff, auszug)
            eindeutig = projekt is not None and kategorie != "Info"
            ki = False

            # --- Stufe 2: nur wenn lokal NICHT eindeutig ---
            if not eindeutig and use_api:
                try:
                    meta = {"absender": absender_smtp, "betreff": betreff, "auszug": auszug}
                    p_kuerzel, k = klassifiziere_stufe2(meta, api_key, projektnamen)
                    if k in KATEGORIEN:
                        kategorie = k
                    if p_kuerzel:
                        treffer = next((p for p in mapping["projekte"]
                                        if p.get("kuerzel", "").lower() == p_kuerzel.lower()
                                        or p.get("name", "").lower() == p_kuerzel.lower()), None)
                        projekt = treffer or projekt or {"name": p_kuerzel, "kuerzel": p_kuerzel}
                    eindeutig = projekt is not None and kategorie != "Info"
                    ki = True
                except Exception as e:  # noqa: BLE001
                    print(f"  Stufe-2-Fehler bei '{betreff[:40]}': {e}")

            relevant = kategorie != "Info"
            gruppe = DRINGLICHKEIT.get(kategorie, "Niedrig")   # rein nach Dringlichkeit
            if relevant and not projekt:
                projekt_anzeige = "(Projekt prüfen)"
            else:
                projekt_anzeige = (projekt or {}).get("name", "—") if projekt else "—"

            zeilen.append({
                "kuerzel": (projekt or {}).get("kuerzel", "") if projekt else "",
                "projekt": projekt_anzeige,
                "kategorie": kategorie,
                "absender": absender_smtp or "(unbekannt)",
                "betreff": betreff,
                "eingang": received.strftime("%d.%m.%Y %H:%M"),
                "gruppe": gruppe,
                "ki": ki,
                "_relevant": relevant,
                "_item": item,
                "_received": received,
                "_entry_id": entry_id,
            })
            if received > neue_max_received:
                neue_max_received = received
        except Exception as e:  # noqa: BLE001
            print(f"  Übersprungen (Lesefehler): {e}")
        item = items.GetNext()

    # --- Übersicht schreiben (immer) ---
    os.makedirs(bericht_dir, exist_ok=True)
    bericht = os.path.join(bericht_dir,
                           f"Triage_{dt.datetime.now():%Y%m%d_%H%M}.html")
    with open(bericht, "w", encoding="utf-8") as f:
        f.write(baue_html(zeilen, args.scharf))
    print(f"\nÜbersicht: {bericht}  ({len(zeilen)} Mail(s))")
    try:
        os.startfile(bericht)   # noqa: PERF203  (nur Windows)
    except Exception:
        pass

    # --- Ablage nur im scharfen Lauf: ALLE Mails uebergeben (alles umbenennen) ---
    if args.scharf:
        os.makedirs(eingang, exist_ok=True)
        for z in zeilen:
            try:
                stamm = f"{z['_received']:%y%m%d}_{z['kuerzel'] or 'X'}_{sicherer_name(z['betreff'])}"
                ziel = os.path.join(eingang, stamm + ".msg")
                i = 2
                while os.path.exists(ziel):
                    ziel = os.path.join(eingang, f"{stamm}-{i:02d}.msg")
                    i += 1
                z["_item"].SaveAs(ziel, 9)   # 9 = olMSGUnicode (Export, read-only)
                n_export += 1
                if args.mit_kategorie:
                    kat = cfg["outlook_kategorie"]
                    vorhandene = z["_item"].Categories or ""
                    if kat not in vorhandene:
                        z["_item"].Categories = (vorhandene + ";" + kat).strip(";")
                        z["_item"].Save()    # einzige erlaubte, reversible Schreibaktion
            except Exception as e:  # noqa: BLE001
                print(f"  Export-Fehler '{z['betreff'][:40]}': {e}")
        # Marker nur im scharfen Lauf weiterstellen
        state["letzte_received"] = neue_max_received.isoformat()
        state["verarbeitete_entry_ids"] = list(verarbeitet | {z["_entry_id"] for z in zeilen})
        speichere_state(state)
        print(f"Scharf: {n_export} Mail(s) als .msg nach {eingang} gelegt "
              "(der Dokumentenbenennungs-Skill extrahiert/benennt sie). "
              "Marker weitergestellt.")
    else:
        print("Trockenlauf: nichts abgelegt, Marker unverändert. "
              "Mit --scharf scharf schalten.")


if __name__ == "__main__":
    main()
