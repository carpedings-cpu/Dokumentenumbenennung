#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPC Morgenbriefing (Outlook, Windows, COM/pywin32).

Liest Posteingang + Gesendete seit dem letzten Lauf (READ-ONLY), fasst die Mails
mit Google Gemini zusammen, pflegt eine fortlaufende AUFGABENLISTE (abhakbar,
erledigte verschwinden dauerhaft), erkennt TERMINE/FRISTEN inkl. Terminaenderungen
und traegt bestaetigte Termine MIT Projekt in den Outlook-Kalender ein.

Unterscheidung "fuer wen": Steht man im AN -> eigene Aufgabe; nur in KOPIE/CC ->
jemand anderes ist zustaendig (separat ausgewiesen).

SICHERHEIT:
  - E-Mails werden nur GELESEN. Einzige Schreibaktion: bestaetigte Kalender-Termine.
  - Inkrementell ueber Marker (briefing_state.json).
  - Fuer die Zusammenfassung gehen Betreff + Textauszug an die Gemini-API.
"""

import argparse
import datetime as dt
import html
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

HIER = (os.path.dirname(sys.executable)
        if getattr(sys, "frozen", False)            # als .exe: Ordner NEBEN der .exe
        else os.path.dirname(os.path.abspath(__file__)))

ENV_VORLAGE = (
    "# KPC Morgenbriefing - Einstellungen\n"
    "# Gemini-Schluessel holen: https://aistudio.google.com/apikey\n"
    "GEMINI_API_KEY=AIza...\n"
)

CONFIG = {
    "base_dir": r"C:\Users\ziegler\Desktop\Dokumentenumbenennung",
    "gesendete_einbeziehen": True,
    "max_mails": 70,
    "stunden_rueckblick_erststart": 24,
    "gemini_modell": "gemini-2.5-flash",
}

GEMINI_ENDPUNKT = "https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent"
GRUPPEN = ["Hoch", "Mittel", "Niedrig"]


def _melde(titel, text, fehler=False):
    print(text)
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        (messagebox.showerror if fehler else messagebox.showinfo)(titel, text)
        r.destroy()
    except Exception:  # noqa: BLE001
        pass


def oeffne_datei(pfad):
    try:
        os.startfile(pfad)   # nur Windows
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        import webbrowser
        webbrowser.open("file:///" + os.path.abspath(pfad).replace("\\", "/"))
    except Exception:  # noqa: BLE001
        pass


# ---------------------------------------------------------------------------
# Konfiguration / .env
# ---------------------------------------------------------------------------
def lade_config():
    pfad = os.path.join(HIER, "briefing_config.json")
    if os.path.exists(pfad):
        try:
            CONFIG.update(json.load(open(pfad, encoding="utf-8")))
        except Exception as e:  # noqa: BLE001
            print(f"WARN: briefing_config.json nicht lesbar: {e}")
    return CONFIG


def lade_env(pfad):
    werte = {}
    if os.path.exists(pfad):
        for zeile in open(pfad, encoding="utf-8-sig"):
            zeile = zeile.strip()
            if not zeile or zeile.startswith("#") or "=" not in zeile:
                continue
            k, v = zeile.split("=", 1)
            werte[k.strip()] = re.sub(r"\s+", "", v).strip('"').strip("'")
    return werte


def gemini_key():
    env = lade_env(os.path.join(HIER, ".env"))
    return (env.get("GEMINI_API_KEY") or env.get("GOOGLE_API_KEY")
            or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or "").strip()


# ---------------------------------------------------------------------------
# Status (Marker + bekannte Termine) und Aufgabenliste
# ---------------------------------------------------------------------------
def state_pfad():
    return os.path.join(HIER, "briefing_state.json")


def lade_state():
    if os.path.exists(state_pfad()):
        try:
            d = json.load(open(state_pfad(), encoding="utf-8"))
            d.setdefault("letzter_lauf", None)
            d.setdefault("termine", {})       # key -> {datum, titel, projekt}
            return d
        except Exception:  # noqa: BLE001
            pass
    return {"letzter_lauf": None, "termine": {}}


def speichere_state(state):
    json.dump(state, open(state_pfad(), "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def todo_pfad():
    return os.path.join(HIER, "todos.json")


def lade_todos():
    if os.path.exists(todo_pfad()):
        try:
            d = json.load(open(todo_pfad(), encoding="utf-8"))
            if isinstance(d, dict) and isinstance(d.get("todos"), list):
                return d
        except Exception:  # noqa: BLE001
            pass
    return {"todos": []}


def speichere_todos(store):
    json.dump(store, open(todo_pfad(), "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def _norm(s):
    return re.sub(r"[^0-9a-zäöüß ]+", " ", (s or "").lower()).strip()


def schluessel(projekt, titel):
    return _norm(projekt) + "|" + _norm(titel)[:70]


def merge_todos(store, punkte, heute):
    """Fuegt neue Aufgaben hinzu; bereits bekannte (offen ODER erledigt) nicht erneut."""
    bekannt = {t["id"] for t in store["todos"]}
    neu = 0
    for p in punkte:
        thema = (p.get("thema") or "").strip()
        if not thema:
            continue
        sid = schluessel(p.get("projekt", ""), thema)
        if sid in bekannt:
            continue
        bekannt.add(sid)
        store["todos"].append({
            "id": sid,
            "projekt": p.get("projekt", "Allgemein"),
            "thema": thema,
            "schritt": p.get("naechster_schritt", ""),
            "dringlichkeit": p.get("dringlichkeit", "Niedrig"),
            "fuer_mich": bool(p.get("fuer_mich", True)),
            "richtung": p.get("richtung", ""),
            "status": "offen",
            "erstellt": heute,
            "erledigt_am": None,
        })
        neu += 1
    return neu


# ---------------------------------------------------------------------------
# Projektliste (zur Gruppierung)
# ---------------------------------------------------------------------------
def lade_projektnamen(cfg):
    kandidaten = [
        os.path.join(cfg["base_dir"], "projekte_mapping.json"),
        os.path.join(HIER, "projekte_mapping.json"),
        os.path.join(HIER, "..", "triage", "projekte_mapping.json"),
    ]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        kandidaten.append(os.path.join(meipass, "projekte_mapping.json"))
    for pfad in kandidaten:
        try:
            if os.path.exists(pfad):
                daten = json.load(open(pfad, encoding="utf-8"))
                projekte = daten.get("projekte", []) if isinstance(daten, dict) else []
                namen = [f"{p.get('kuerzel', '')} - {p.get('name', '')}".strip(" -")
                         for p in projekte if p.get("name")]
                if namen:
                    return namen
        except Exception:  # noqa: BLE001
            continue
    return []


# ---------------------------------------------------------------------------
# Outlook (READ-ONLY lesen)
# ---------------------------------------------------------------------------
def outlook_app():
    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError("pywin32 wird benoetigt: pip install pywin32") from e
    return win32com.client.Dispatch("Outlook.Application")


def _ordner(app, folder_id):
    return app.GetNamespace("MAPI").GetDefaultFolder(folder_id)


def py_datetime(com_zeit):
    try:
        return dt.datetime(com_zeit.year, com_zeit.month, com_zeit.day,
                           com_zeit.hour, com_zeit.minute, com_zeit.second)
    except Exception:  # noqa: BLE001
        return dt.datetime.fromtimestamp(0)


def smtp_adresse(item):
    try:
        if (item.SenderEmailType or "").upper() == "EX":
            try:
                return item.Sender.GetExchangeUser().PrimarySmtpAddress or ""
            except Exception:  # noqa: BLE001
                PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                try:
                    return item.PropertyAccessor.GetProperty(PR_SMTP) or ""
                except Exception:  # noqa: BLE001
                    return item.SenderEmailAddress or ""
        return item.SenderEmailAddress or ""
    except Exception:  # noqa: BLE001
        return ""


def empfaenger_anzeige(item):
    try:
        return item.To or "(unbekannt)"
    except Exception:  # noqa: BLE001
        return "(unbekannt)"


def eigene_adressen(app):
    """SMTP-Adressen des eigenen Postfachs (fuer An/Kopie-Unterscheidung)."""
    adr = set()
    try:
        ns = app.GetNamespace("MAPI")
        try:
            cu = ns.CurrentUser
            try:
                a = cu.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                if a:
                    adr.add(a.lower())
            except Exception:  # noqa: BLE001
                pass
            try:
                if cu.Address and "@" in cu.Address:
                    adr.add(cu.Address.lower())
            except Exception:  # noqa: BLE001
                pass
        except Exception:  # noqa: BLE001
            pass
        try:
            for acc in ns.Accounts:
                try:
                    if acc.SmtpAddress:
                        adr.add(acc.SmtpAddress.lower())
                except Exception:  # noqa: BLE001
                    pass
        except Exception:  # noqa: BLE001
            pass
    except Exception:  # noqa: BLE001
        pass
    return adr


def _recipient_smtp(r):
    try:
        ae = r.AddressEntry
        if (ae.Type or "").upper() == "EX":
            try:
                return (ae.GetExchangeUser().PrimarySmtpAddress or "").lower()
            except Exception:  # noqa: BLE001
                pass
        return (r.Address or "").lower()
    except Exception:  # noqa: BLE001
        return (getattr(r, "Address", "") or "").lower()


def rolle_eingang(item, eigene):
    """'Mich' (im An), 'Kopie' (nur CC) oder 'Mich' als Default, wenn unklar."""
    if not eigene:
        return "Mich"
    in_to = in_cc = False
    try:
        rec = item.Recipients
        anzahl = min(int(rec.Count), 60)
        for i in range(1, anzahl + 1):
            r = rec.Item(i)
            smtp = _recipient_smtp(r)
            if smtp and smtp in eigene:
                typ = int(getattr(r, "Type", 1))
                if typ == 1:
                    in_to = True
                elif typ == 2:
                    in_cc = True
    except Exception:  # noqa: BLE001
        return "Mich"
    if in_to:
        return "Mich"
    if in_cc:
        return "Kopie"
    return "Mich"


def sammle_mails(cfg, marker, max_n):
    app = outlook_app()
    eigene = eigene_adressen(app)
    quellen = [(6, "Eingang")]
    if cfg.get("gesendete_einbeziehen", True):
        quellen.append((5, "Gesendet"))

    mails = []
    for fid, richtung in quellen:
        try:
            ordner = _ordner(app, fid)
        except Exception as e:  # noqa: BLE001
            print(f"  Ordner {richtung} nicht verfuegbar: {e}")
            continue
        items = ordner.Items
        try:
            items.Sort("[SentOn]" if richtung == "Gesendet" else "[ReceivedTime]", True)
        except Exception:  # noqa: BLE001
            items.Sort("[ReceivedTime]", True)

        item = items.GetFirst()
        gezaehlt = 0
        while item is not None and gezaehlt < max_n:
            try:
                if int(getattr(item, "Class", 0)) != 43:   # 43 = olMail
                    item = items.GetNext(); continue
                if richtung == "Gesendet":
                    zeit = py_datetime(getattr(item, "SentOn", None) or item.ReceivedTime)
                else:
                    zeit = py_datetime(item.ReceivedTime)
                if zeit <= marker:
                    break
                betreff = item.Subject or ""
                if richtung == "Gesendet":
                    partner = "An: " + empfaenger_anzeige(item)
                    rolle = "VonMir"
                else:
                    partner = smtp_adresse(item) or "(unbekannt)"
                    rolle = rolle_eingang(item, eigene)
                try:
                    auszug = re.sub(r"[ \t]+", " ", (item.Body or "")).strip()[:1200]
                except Exception:  # noqa: BLE001
                    auszug = ""
                mails.append({"richtung": richtung, "rolle": rolle, "zeit": zeit,
                              "partner": partner, "betreff": betreff, "auszug": auszug})
                gezaehlt += 1
            except Exception as e:  # noqa: BLE001
                print(f"  Uebersprungen (Lesefehler): {e}")
            item = items.GetNext()
    mails.sort(key=lambda m: m["zeit"], reverse=True)
    return mails


WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]


def lese_kalender_woche(jetzt):
    """Outlook-Kalender READ-ONLY: Termine von heute bis Sonntag dieser Woche."""
    heute = jetzt.replace(hour=0, minute=0, second=0, microsecond=0)
    montag = heute - dt.timedelta(days=heute.weekday())
    ende = montag + dt.timedelta(days=6, hours=23, minutes=59)
    ergebnis = []
    try:
        app = outlook_app()
        cal = app.GetNamespace("MAPI").GetDefaultFolder(9)   # 9 = olFolderCalendar
        items = cal.Items
        try:
            items.Sort("[Start]")
            items.IncludeRecurrences = True
        except Exception:  # noqa: BLE001
            pass
        it = items.GetFirst()
        zaehler = 0
        while it is not None and zaehler < 1000:
            zaehler += 1
            try:
                if int(getattr(it, "Class", 0)) == 26:   # 26 = olAppointment
                    s = py_datetime(it.Start)
                    if s > ende:
                        break                              # aufsteigend sortiert -> fertig
                    if s >= heute:
                        ganztags = bool(getattr(it, "AllDayEvent", False))
                        ergebnis.append({
                            "sort": s,
                            "wochentag": WOCHENTAGE[s.weekday()],
                            "datum": s.strftime("%d.%m."),
                            "uhrzeit": "" if ganztags else s.strftime("%H:%M"),
                            "titel": it.Subject or "",
                            "ort": getattr(it, "Location", "") or "",
                        })
            except Exception:  # noqa: BLE001
                pass
            it = items.GetNext()
    except Exception as e:  # noqa: BLE001
        print(f"  Kalender nicht lesbar: {e}")
    ergebnis.sort(key=lambda x: x["sort"])
    return ergebnis


# ---------------------------------------------------------------------------
# Gemini
# ---------------------------------------------------------------------------
def gemini_json(api_key, system_text, user_text, modell):
    body = {
        "contents": [{"parts": [{"text": user_text}]}],
        "system_instruction": {"parts": [{"text": system_text}]},
        "generationConfig": {"response_mime_type": "application/json", "temperature": 0.2},
    }
    url = GEMINI_ENDPUNKT.format(m=modell) + "?key=" + urllib.parse.quote(api_key)
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            daten = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        text = e.read().decode("utf-8", "replace")
        raise RuntimeError(f"Gemini-Fehler {e.code}: {text[:300]}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Netzwerkfehler zur Gemini-API: {e}") from e
    try:
        cand = daten["candidates"][0]
        text = "".join(p.get("text", "") for p in cand["content"]["parts"])
        return json.loads(text)
    except (KeyError, IndexError, ValueError) as e:
        raise RuntimeError(f"Unerwartete Gemini-Antwort: {json.dumps(daten)[:300]}") from e


_ROLLE_TAG = {"Mich": "[AN MICH]", "Kopie": "[NUR KOPIE - jemand anderes zustaendig]",
              "VonMir": "[VON MIR gesendet]"}


def baue_briefing(mails, projektnamen, cfg, api_key):
    heute = dt.date.today().isoformat()
    system = (
        "Du bist Assistent fuer einen Bauprojektleiter der Firma KPC. Erstelle aus "
        "den E-Mails ein knappes, sachliches Morgenbriefing auf Deutsch. "
        f"Heutiges Datum: {heute}. Loese relative Angaben (heute, morgen, Freitag, "
        "naechste Woche) anhand des heutigen Datums auf.\n"
        "Jede Mail ist markiert: [AN MICH] = an mich gerichtet (meine Aufgabe); "
        "[NUR KOPIE ...] = ich bin nur in Kopie, jemand anderes ist zustaendig; "
        "[VON MIR gesendet] = von mir verschickt (Nachfassen/Warten auf Antwort).\n"
        "Regeln:\n"
        "- Ordne jeden Punkt einem Projekt aus der Liste zu (nur den Namen); passt "
        "keins, schreibe 'Allgemein'.\n"
        "- fuer_mich: true wenn die Aufgabe MEINE ist ([AN MICH] oder [VON MIR]); "
        "false wenn ich nur in Kopie bin ([NUR KOPIE]).\n"
        "- dringlichkeit: 'Hoch' (Maengel, Behinderung, Fristen, Eskalation, Termin "
        "heute/morgen), 'Mittel' (Rechnung, Lieferavis, Antwort noetig), 'Niedrig' (Info).\n"
        "- termine NUR bei konkretem Datum/Frist. datum als YYYY-MM-DD; uhrzeit 'HH:MM' "
        "oder leer (ganztaegig); dauer_min Standard 60; projekt dazuschreiben.\n"
        "- verschoben: true, wenn die Mail einen BESTEHENDEN Termin aendert/verlegt; "
        "dann alt_datum (YYYY-MM-DD) wenn erkennbar, sonst leer.\n"
        "- Fasse zusammen, erfinde nichts. Antworte AUSSCHLIESSLICH als JSON nach diesem Schema:\n"
        '{"ueberblick": "2-4 Saetze", '
        '"punkte": [{"projekt": "", "dringlichkeit": "Hoch|Mittel|Niedrig", '
        '"richtung": "Eingang|Gesendet", "fuer_mich": true, "thema": "", '
        '"naechster_schritt": ""}], '
        '"termine": [{"projekt": "", "titel": "", "datum": "YYYY-MM-DD", "uhrzeit": "", '
        '"dauer_min": 60, "ort": "", "quelle": "", "verschoben": false, "alt_datum": ""}]}'
    )
    zeilen = []
    if projektnamen:
        zeilen.append("Bekannte Projekte:\n" + "\n".join(projektnamen) + "\n")
    zeilen.append(f"E-Mails ({len(mails)}):")
    for i, m in enumerate(mails, 1):
        tag = _ROLLE_TAG.get(m.get("rolle", "Mich"), "")
        zeilen.append(
            f"[{i}] {tag} ({m['richtung']}) {m['zeit']:%Y-%m-%d %H:%M} | {m['partner']} | "
            f"Betreff: {m['betreff']}\nAuszug: {m['auszug']}")
    daten = gemini_json(api_key, system, "\n".join(zeilen), cfg["gemini_modell"])
    daten.setdefault("ueberblick", "")
    daten.setdefault("punkte", [])
    daten.setdefault("termine", [])
    return daten


# ---------------------------------------------------------------------------
# Terminaenderungen erkennen
# ---------------------------------------------------------------------------
def markiere_terminaenderungen(termine, state):
    """Markiert Termine als 'geaendert', wenn ein bekannter Termin (Projekt+Titel)
    jetzt ein anderes Datum hat - oder die Mail selbst eine Verschiebung meldet."""
    bekannt = state.get("termine", {})
    for t in termine:
        sid = schluessel(t.get("projekt", ""), t.get("titel", ""))
        t["_key"] = sid
        alt = bekannt.get(sid)
        geaendert = False
        alt_datum = t.get("alt_datum", "") or ""
        if alt and alt.get("datum") and alt["datum"] != t.get("datum", ""):
            geaendert = True
            alt_datum = alt_datum or alt["datum"]
        if t.get("verschoben"):
            geaendert = True
        t["geaendert"] = geaendert
        t["alt_datum"] = alt_datum
    return termine


# ---------------------------------------------------------------------------
# HTML-Briefing
# ---------------------------------------------------------------------------
def _termin_label(t):
    txt = f"{t.get('datum', '')}  {t.get('uhrzeit', '') or '(ganztags)'}"
    if t.get("geaendert"):
        alt = t.get("alt_datum", "")
        txt = ("GEAENDERT: " + (f"war {alt} -> " if alt else "") + txt)
    return txt


def baue_html(ueberblick, offene, termine, wochentermine, zeitraum, anzahl_mails):
    def esc(x):
        return html.escape(str(x or ""))

    meine = [t for t in offene if t.get("fuer_mich", True)]
    info = [t for t in offene if not t.get("fuer_mich", True)]
    nach_gruppe = {g: [] for g in GRUPPEN}
    for t in meine:
        nach_gruppe.get(t.get("dringlichkeit", "Niedrig"), nach_gruppe["Niedrig"]).append(t)
    stand = dt.datetime.now().strftime("%A, %d.%m.%Y %H:%M")

    teile = ["""<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">
<title>KPC Morgenbriefing</title>
<link href="https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@300;400;700&display=swap" rel="stylesheet">
<style>
  @page { size: A4; margin: 14mm; }
  * { box-sizing: border-box; }
  body { font-family:'Roboto Condensed','Arial Narrow',sans-serif; font-weight:300;
         color:#2d2926; margin:0; padding:28px; background:#fff; }
  header { border-bottom:3px solid #c8a882; padding-bottom:10px; margin-bottom:16px; }
  h1 { font-weight:300; font-size:26px; letter-spacing:1px; margin:0; }
  .meta { color:#6b5040; font-size:12px; margin-top:4px; }
  .ueberblick { background:#faf6ef; border-left:6px solid #c8a882; padding:12px 14px;
                margin:14px 0; font-size:14px; }
  h2 { font-weight:400; font-size:15px; color:#fff; background:#2d2926;
       padding:6px 10px; margin:18px 0 0 0; border-left:6px solid #c8a882; }
  h2.mittel { background:#6b5040; } h2.niedrig { background:#9a8a78; }
  h2.info { background:#9a8a78; } h2.termine { background:#3c5a4a; }
  table { width:100%; border-collapse:collapse; font-size:12px; }
  th { text-align:left; background:#6b5040; color:#fff; font-weight:400; padding:6px 8px; }
  td { padding:6px 8px; border-bottom:1px solid #e6ddcf; vertical-align:top; }
  tr:nth-child(even) td { background:#faf6ef; }
  .proj { color:#6b5040; font-weight:700; }
  .chg { color:#8a2f2f; font-weight:700; }
  footer { margin-top:22px; color:#6b5040; font-size:11px;
           border-top:1px solid #c8a882; padding-top:8px; }
</style></head><body>
<header><h1>KPC &middot; Morgenbriefing</h1>
<div class="meta">Stand: """ + esc(stand) + " &nbsp;|&nbsp; Zeitraum: " + esc(zeitraum)
        + " &nbsp;|&nbsp; Neue Mails: " + str(anzahl_mails)
        + " &nbsp;|&nbsp; Offene Aufgaben: " + str(len(offene)) + """</div></header>"""]

    if ueberblick:
        teile.append('<div class="ueberblick">' + esc(ueberblick) + "</div>")

    if wochentermine:
        teile.append(f'<h2 class="termine">Termine diese Woche ({len(wochentermine)})</h2>')
        teile.append("<table><tr><th>Tag</th><th>Datum</th><th>Uhrzeit</th>"
                     "<th>Titel</th><th>Ort</th></tr>")
        for t in wochentermine:
            teile.append("<tr>"
                         f"<td>{esc(t['wochentag'])}</td>"
                         f"<td>{esc(t['datum'])}</td>"
                         f"<td>{esc(t['uhrzeit'] or 'ganztags')}</td>"
                         f"<td>{esc(t['titel'])}</td>"
                         f"<td>{esc(t['ort'])}</td></tr>")
        teile.append("</table>")

    klasse = {"Hoch": "", "Mittel": "mittel", "Niedrig": "niedrig"}
    titel = {"Hoch": "Meine Aufgaben - wichtig / dringend",
             "Mittel": "Meine Aufgaben - zu erledigen",
             "Niedrig": "Meine Aufgaben - nachrangig"}
    for g in GRUPPEN:
        rows = nach_gruppe[g]
        if not rows:
            continue
        teile.append(f'<h2 class="{klasse[g]}">{esc(titel[g])} ({len(rows)})</h2>')
        teile.append("<table><tr><th>Projekt</th><th>Thema</th>"
                     "<th>Naechster Schritt</th></tr>")
        for t in rows:
            teile.append("<tr>"
                         f'<td><span class="proj">{esc(t.get("projekt", "Allgemein"))}</span></td>'
                         f"<td>{esc(t.get('thema', ''))}</td>"
                         f"<td>{esc(t.get('schritt', ''))}</td></tr>")
        teile.append("</table>")

    if info:
        teile.append(f'<h2 class="info">Nur zur Info - jemand anderes zustaendig ({len(info)})</h2>')
        teile.append("<table><tr><th>Projekt</th><th>Thema</th><th>Hinweis</th></tr>")
        for t in info:
            teile.append("<tr>"
                         f'<td><span class="proj">{esc(t.get("projekt", "Allgemein"))}</span></td>'
                         f"<td>{esc(t.get('thema', ''))}</td>"
                         f"<td>{esc(t.get('schritt', ''))}</td></tr>")
        teile.append("</table>")

    if termine:
        teile.append(f'<h2 class="termine">Neue Termine aus Mails - in Kalender uebernehmen ({len(termine)})</h2>')
        teile.append("<table><tr><th>Projekt</th><th>Datum</th><th>Uhrzeit</th>"
                     "<th>Titel</th><th>Status</th></tr>")
        for t in termine:
            status = '<span class="chg">GEAENDERT</span>' if t.get("geaendert") else "neu"
            alt = f" (war {esc(t.get('alt_datum'))})" if t.get("geaendert") and t.get("alt_datum") else ""
            teile.append("<tr>"
                         f'<td><span class="proj">{esc(t.get("projekt", ""))}</span></td>'
                         f"<td>{esc(t.get('datum', ''))}</td>"
                         f"<td>{esc(t.get('uhrzeit', '') or 'ganztags')}</td>"
                         f"<td>{esc(t.get('titel', ''))}</td>"
                         f"<td>{status}{alt}</td></tr>")
        teile.append("</table>")

    teile.append('<footer>Aus Outlook (Eingang + Gesendet), Zusammenfassung durch '
                 'Gemini. Mails nur gelesen; Termine nur nach Bestaetigung im Kalender. '
                 'Erledigte Aufgaben werden gemerkt und nicht erneut gezeigt.</footer>'
                 "</body></html>")
    return "".join(teile)


def schreibe_html(ueberblick, offene, termine, wochentermine, zeitraum, anzahl):
    ordner = os.path.join(HIER, "Briefings")
    os.makedirs(ordner, exist_ok=True)
    pfad = os.path.join(ordner, f"Briefing_{dt.datetime.now():%Y%m%d_%H%M}.html")
    with open(pfad, "w", encoding="utf-8") as f:
        f.write(baue_html(ueberblick, offene, termine, wochentermine, zeitraum, anzahl))
    return pfad


def schreibe_log(ueberblick, offene, termine):
    pfad = os.path.join(HIER, "briefing_log.md")
    with open(pfad, "a", encoding="utf-8") as f:
        f.write(f"\n\n## {dt.datetime.now():%Y-%m-%d %H:%M}\n\n")
        if ueberblick:
            f.write(ueberblick + "\n\n")
        for t in offene:
            wer = "MIR" if t.get("fuer_mich", True) else "INFO"
            f.write(f"- [{t.get('dringlichkeit', '')}/{wer}] {t.get('projekt', '')}: "
                    f"{t.get('thema', '')} -> {t.get('schritt', '')}\n")
        for t in termine:
            f.write(f"- TERMIN {t.get('datum', '')} {t.get('uhrzeit', '')} "
                    f"{t.get('projekt', '')} - {t.get('titel', '')}"
                    f"{' (GEAENDERT)' if t.get('geaendert') else ''}\n")
    return pfad


# ---------------------------------------------------------------------------
# Fenster: Aufgaben abhaken
# ---------------------------------------------------------------------------
def abhaken_fenster(offene):
    """Zeigt offene Aufgaben; gibt die IDs der als erledigt markierten zurueck."""
    if not offene:
        return set()
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:  # noqa: BLE001
        return set()
    root = tk.Tk()
    root.title("Aufgaben abhaken - erledigte verschwinden dauerhaft")
    root.geometry("1000x520")
    ttk.Label(root, padding=8, text=("Zeile anklicken = als ERLEDIGT markieren. "
              "Erledigte werden gemerkt und nicht mehr gezeigt.")).pack(anchor="w")
    rahmen = ttk.Frame(root, padding=(8, 0))
    rahmen.pack(fill="both", expand=True)
    cols = ("sel", "dringl", "fuer", "projekt", "thema")
    tree = ttk.Treeview(rahmen, columns=cols, show="headings", height=18)
    for c, t, w in (("sel", "erledigt", 70), ("dringl", "Dringl.", 70),
                    ("fuer", "Fuer", 90), ("projekt", "Projekt", 220),
                    ("thema", "Thema", 520)):
        tree.heading(c, text=t)
        tree.column(c, width=w, anchor="w")
    sb = ttk.Scrollbar(rahmen, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    checked = {}
    row_of = {}
    for t in offene:
        iid = tree.insert("", "end", values=(
            "☐", t.get("dringlichkeit", ""),
            "Mir" if t.get("fuer_mich", True) else "Info",
            t.get("projekt", ""), t.get("thema", "")))
        checked[iid] = False
        row_of[iid] = t

    def klick(event):
        if tree.identify("region", event.x, event.y) != "cell":
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        checked[iid] = not checked[iid]
        vals = list(tree.item(iid, "values"))
        vals[0] = "☑" if checked[iid] else "☐"
        tree.item(iid, values=vals)
    tree.bind("<Button-1>", klick)

    erg = {"ids": set()}
    leiste = ttk.Frame(root, padding=8)
    leiste.pack(fill="x")

    def speichern():
        erg["ids"] = {row_of[i]["id"] for i in tree.get_children() if checked.get(i)}
        root.destroy()
    ttk.Button(leiste, text="Erledigte speichern", command=speichern).pack(side="right")
    ttk.Button(leiste, text="Nichts abhaken", command=root.destroy).pack(side="right", padx=6)
    root.mainloop()
    return erg["ids"]


# ---------------------------------------------------------------------------
# Fenster: Termine bestaetigen + in Outlook-Kalender eintragen
# ---------------------------------------------------------------------------
def bestaetige_termine(termine):
    if not termine:
        return []
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:  # noqa: BLE001
        return []
    root = tk.Tk()
    root.title("Termine in den Outlook-Kalender uebernehmen")
    root.geometry("860x480")
    ttk.Label(root, padding=8, text=("Haken setzen bei den Terminen, die in deinen "
              "Outlook-Kalender sollen (das Projekt steht im Betreff):")).pack(anchor="w")
    rahmen = ttk.Frame(root, padding=(10, 0))
    rahmen.pack(fill="both", expand=True)
    eintraege = []
    for t in termine:
        v = tk.BooleanVar(value=True)
        proj = t.get("projekt", "") or "ohne Projekt"
        txt = f"[{proj}]  {_termin_label(t)}  -  {t.get('titel', '')}"
        ttk.Checkbutton(rahmen, text=txt, variable=v).pack(anchor="w", pady=1)
        eintraege.append((v, t))
    erg = {"ok": False}
    leiste = ttk.Frame(root, padding=8)
    leiste.pack(fill="x")

    def uebernehmen():
        erg["ok"] = True
        root.destroy()
    ttk.Button(leiste, text="Ausgewaehlte in Kalender eintragen",
               command=uebernehmen).pack(side="right")
    ttk.Button(leiste, text="Keine / Schliessen", command=root.destroy).pack(side="right", padx=6)
    root.mainloop()
    if not erg["ok"]:
        return []
    return [t for v, t in eintraege if v.get()]


def _start_zeit(datum, uhrzeit):
    d = dt.datetime.strptime(datum, "%Y-%m-%d")
    if uhrzeit:
        try:
            hh, mm = (uhrzeit.split(":") + ["0"])[:2]
            return d.replace(hour=int(hh), minute=int(mm)), True
        except Exception:  # noqa: BLE001
            return d, False
    return d, False


def trage_termine_ein(termine, state):
    """Legt Termine im Kalender an (Betreff mit Projekt) und merkt sie im Status."""
    if not termine:
        return 0
    import pywintypes
    app = outlook_app()
    n = 0
    for t in termine:
        try:
            start, hat_uhrzeit = _start_zeit(t["datum"], t.get("uhrzeit", ""))
            proj = (t.get("projekt", "") or "").strip()
            titel = t.get("titel", "Termin")
            appt = app.CreateItem(1)   # 1 = olAppointmentItem
            appt.Subject = (f"{proj} - {titel}" if proj else titel)
            appt.Start = pywintypes.Time(start)
            if hat_uhrzeit:
                appt.Duration = int(t.get("dauer_min") or 60)
            else:
                appt.AllDayEvent = True
            if t.get("ort"):
                appt.Location = t["ort"]
            hinweis = ""
            if t.get("geaendert"):
                hinweis = ("\nACHTUNG: geaenderter Termin"
                           + (f" (war {t.get('alt_datum')})" if t.get("alt_datum") else "")
                           + " - bitte alten Kalendereintrag pruefen/loeschen.")
            appt.Body = (f"Projekt: {proj}\nAus KPC-Morgenbriefing.\n"
                         f"Quelle: {t.get('quelle', '') or ''}{hinweis}")
            appt.ReminderSet = True
            appt.Save()
            n += 1
            key = t.get("_key") or schluessel(proj, titel)
            state.setdefault("termine", {})[key] = {
                "datum": t.get("datum", ""), "titel": titel, "projekt": proj}
        except Exception as e:  # noqa: BLE001
            print(f"  Termin '{t.get('titel', '')}' nicht eingetragen: {e}")
    return n


# ---------------------------------------------------------------------------
# Hauptlauf
# ---------------------------------------------------------------------------
def _marker_normal(state, cfg, jetzt):
    if state.get("letzter_lauf"):
        return dt.datetime.fromisoformat(state["letzter_lauf"])
    return jetzt - dt.timedelta(hours=cfg["stunden_rueckblick_erststart"])


def waehle_zeitraum(state, cfg, jetzt):
    """Kleines Startfenster: welchen Zeitraum auswerten? Gibt den Marker zurueck."""
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:  # noqa: BLE001
        return _marker_normal(state, cfg, jetzt)
    wahl = {"v": "normal"}
    root = tk.Tk()
    root.title("Morgenbriefing - Zeitraum")
    root.geometry("460x300")
    ttk.Label(root, padding=12, font=("", 11, "bold"),
              text="Welchen Zeitraum soll das Briefing auswerten?").pack(anchor="w")
    ttk.Label(root, padding=(12, 0), foreground="#6b5040",
              text=("Normal = nur neue Mails seit dem letzten Lauf.\n"
                    "Weiter zurueck holt aeltere Aufgaben erneut "
                    "(Doppelte werden vermieden).")).pack(anchor="w")

    def setze(v):
        wahl["v"] = v
        root.destroy()
    f = ttk.Frame(root, padding=12)
    f.pack(fill="x")
    ttk.Button(f, text="Seit letztem Briefing (normal)",
               command=lambda: setze("normal")).pack(fill="x", pady=3)
    ttk.Button(f, text="Heute (ab 0:00 Uhr)",
               command=lambda: setze("heute")).pack(fill="x", pady=3)
    ttk.Button(f, text="Letzte 3 Tage",
               command=lambda: setze("3")).pack(fill="x", pady=3)
    ttk.Button(f, text="Letzte 7 Tage",
               command=lambda: setze("7")).pack(fill="x", pady=3)
    root.mainloop()
    v = wahl["v"]
    if v == "heute":
        return jetzt.replace(hour=0, minute=0, second=0, microsecond=0)
    if v == "3":
        return jetzt - dt.timedelta(days=3)
    if v == "7":
        return jetzt - dt.timedelta(days=7)
    return _marker_normal(state, cfg, jetzt)


def _lauf(args):
    cfg = lade_config()
    key = gemini_key()
    if not key:
        pfad = os.path.join(HIER, ".env")
        if not os.path.exists(pfad):
            try:
                with open(pfad, "w", encoding="utf-8") as f:
                    f.write(ENV_VORLAGE)
            except Exception:  # noqa: BLE001
                pass
        _melde("Schluessel fehlt",
               "Es ist noch kein Gemini-Schluessel hinterlegt.\n\n"
               f"Bitte die Datei .env (neben dem Programm:\n{pfad})\n"
               "oeffnen und eintragen:\n\n  GEMINI_API_KEY=AIza...\n\n"
               "Schluessel holen: https://aistudio.google.com/apikey\n"
               "Danach erneut starten.", fehler=True)
        return

    state = lade_state()
    store = lade_todos()
    jetzt = dt.datetime.now()
    heute = jetzt.date().isoformat()
    if args.seit:
        marker = dt.datetime.strptime(args.seit, "%Y-%m-%d")
    elif args.stunden:
        marker = jetzt - dt.timedelta(hours=args.stunden)
    elif args.auto:
        marker = _marker_normal(state, cfg, jetzt)     # ohne Nachfrage (Aufgabenplanung)
    else:
        marker = waehle_zeitraum(state, cfg, jetzt)

    print(f"Lese Outlook ab {marker:%d.%m.%Y %H:%M} ...")
    mails = sammle_mails(cfg, marker, args.max or cfg["max_mails"])

    ueberblick, termine = "", []
    if mails:
        print(f"{len(mails)} Mail(s). Erstelle Briefing mit Gemini ...")
        brief = baue_briefing(mails, lade_projektnamen(cfg), cfg, key)
        ueberblick = brief.get("ueberblick", "")
        merge_todos(store, brief.get("punkte", []), heute)
        termine = markiere_terminaenderungen(brief.get("termine", []), state)
    else:
        ueberblick = "Keine neuen Mails seit dem letzten Lauf."

    print("Lese Kalender (Wochenvorschau) ...")
    wochentermine = lese_kalender_woche(jetzt)

    offene = [t for t in store["todos"] if t["status"] == "offen"]
    if not mails and not offene and not wochentermine:
        _melde("Nichts Neues",
               "Keine neuen Mails, keine offenen Aufgaben und keine Termine diese Woche.")
        return

    # Aufgaben abhaken
    if not args.kein_abhaken:
        erledigt = abhaken_fenster(offene)
        if erledigt:
            for t in store["todos"]:
                if t["id"] in erledigt and t["status"] == "offen":
                    t["status"] = "erledigt"
                    t["erledigt_am"] = heute
    speichere_todos(store)
    offene = [t for t in store["todos"] if t["status"] == "offen"]

    zeitraum = f"{marker:%d.%m.%Y %H:%M} - {jetzt:%d.%m.%Y %H:%M}"
    html_pfad = schreibe_html(ueberblick, offene, termine, wochentermine, zeitraum, len(mails))
    schreibe_log(ueberblick, offene, termine)
    print(f"Briefing: {html_pfad}")
    oeffne_datei(html_pfad)

    # Termine bestaetigen + eintragen
    n = 0
    if termine and not args.kein_kalender:
        auswahl = bestaetige_termine(termine)
        n = trage_termine_ein(auswahl, state)

    if mails:
        state["letzter_lauf"] = jetzt.isoformat()
    speichere_state(state)

    _melde("Morgenbriefing fertig",
           f"{len(mails)} neue Mail(s) ausgewertet.\n"
           f"{len(offene)} offene Aufgabe(n).\n"
           f"{n} Termin(e) in den Kalender eingetragen.\n\n"
           f"Briefing:\n{html_pfad}")


def main():
    ap = argparse.ArgumentParser(description="KPC Morgenbriefing (Outlook, read-only).")
    ap.add_argument("--seit", default=None, help="Startdatum YYYY-MM-DD.")
    ap.add_argument("--stunden", type=int, default=0, help="Rueckblick in Stunden.")
    ap.add_argument("--kein-kalender", action="store_true", help="Kein Termin-Fenster.")
    ap.add_argument("--kein-abhaken", action="store_true", help="Kein Aufgaben-Fenster.")
    ap.add_argument("--auto", action="store_true",
                    help="Ohne Zeitraum-Abfrage starten (fuer die Aufgabenplanung).")
    ap.add_argument("--max", type=int, default=0, help="Maximale Mailanzahl (Debug).")
    args, _ = ap.parse_known_args()
    try:
        _lauf(args)
    except Exception as e:  # noqa: BLE001
        _melde("Fehler", f"Das Briefing konnte nicht erstellt werden:\n\n{e}\n\n"
               "Tipp: Ist das klassische Outlook geoeffnet und der Schluessel in der "
               ".env korrekt?", fehler=True)


if __name__ == "__main__":
    main()
