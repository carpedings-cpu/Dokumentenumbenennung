#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPC Morgenbriefing (Outlook, Windows, COM/pywin32).

Liest Posteingang + Gesendete seit dem letzten Briefing (READ-ONLY), fasst die
Mails mit Google Gemini zu einem Briefing zusammen, erkennt Termine/Fristen,
schreibt eine HTML-Uebersicht (KPC-Design, oeffnet sie automatisch) und ein
laufendes Protokoll (das "Gedaechtnis"). Erkannte Termine werden in einem
Fenster zur Auswahl angeboten und NUR nach Bestaetigung in den Outlook-Kalender
eingetragen.

SICHERHEIT:
  - E-Mails werden ausschliesslich GELESEN. Einzige Schreibaktion ist das
    Anlegen der vom Nutzer bestaetigten Kalender-Termine.
  - Inkrementell: nur Mails seit dem letzten Lauf (Marker in briefing_state.json).
  - Fuer die Zusammenfassung gehen Betreff + Textauszug an die Gemini-API
    (vom Nutzer ausdruecklich gewuenscht).

Aufruf:
  python briefing.py                 # Briefing seit letztem Lauf
  python briefing.py --stunden 48    # Rueckblick 48 Stunden
  python briefing.py --seit 2026-06-20
  python briefing.py --kein-kalender # nur Briefing, kein Termin-Fenster
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

HIER = os.path.dirname(os.path.abspath(__file__))

CONFIG = {
    "base_dir": r"C:\Users\ziegler\Desktop\Dokumentenumbenennung",
    "briefing_unterordner": "Morgenbriefing",
    "gesendete_einbeziehen": True,
    "max_mails": 70,                  # Schutz gegen riesige API-Anfragen
    "stunden_rueckblick_erststart": 24,
    "gemini_modell": "gemini-2.5-flash",
}

GEMINI_ENDPUNKT = "https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent"
GRUPPEN = ["Hoch", "Mittel", "Niedrig"]


# ---------------------------------------------------------------------------
# Konfiguration / .env / Status
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
        for zeile in open(pfad, encoding="utf-8"):
            zeile = zeile.strip()
            if not zeile or zeile.startswith("#") or "=" not in zeile:
                continue
            k, v = zeile.split("=", 1)
            werte[k.strip()] = v.strip().strip('"').strip("'")
    return werte


def gemini_key():
    env = lade_env(os.path.join(HIER, ".env"))
    return (env.get("GEMINI_API_KEY") or env.get("GOOGLE_API_KEY")
            or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or "").strip()


def state_pfad():
    return os.path.join(HIER, "briefing_state.json")


def lade_state():
    if os.path.exists(state_pfad()):
        try:
            return json.load(open(state_pfad(), encoding="utf-8"))
        except Exception:  # noqa: BLE001
            pass
    return {"letzter_lauf": None}


def speichere_state(state):
    json.dump(state, open(state_pfad(), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Projektliste (zur Gruppierung) - aus projekte_mapping.json, falls vorhanden
# ---------------------------------------------------------------------------
def lade_projektnamen(cfg):
    kandidaten = [
        os.path.join(cfg["base_dir"], "projekte_mapping.json"),
        os.path.join(HIER, "projekte_mapping.json"),
        os.path.join(HIER, "..", "triage", "projekte_mapping.json"),
    ]
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
# Outlook (READ-ONLY)
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


def sammle_mails(cfg, marker, max_n):
    """Liest Posteingang (und optional Gesendete) READ-ONLY ab dem Marker."""
    app = outlook_app()
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
                else:
                    partner = smtp_adresse(item) or "(unbekannt)"
                try:
                    auszug = re.sub(r"[ \t]+", " ", (item.Body or "")).strip()[:1200]
                except Exception:  # noqa: BLE001
                    auszug = ""
                mails.append({
                    "richtung": richtung,
                    "zeit": zeit,
                    "partner": partner,
                    "betreff": betreff,
                    "auszug": auszug,
                })
                gezaehlt += 1
            except Exception as e:  # noqa: BLE001
                print(f"  Uebersprungen (Lesefehler): {e}")
            item = items.GetNext()
    mails.sort(key=lambda m: m["zeit"], reverse=True)
    return mails


# ---------------------------------------------------------------------------
# Gemini (Zusammenfassung + Terminerkennung)
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


def baue_briefing(mails, projektnamen, cfg, api_key):
    heute = dt.date.today().isoformat()
    system = (
        "Du bist Assistent fuer einen Bauprojektleiter der Firma KPC. Erstelle aus "
        "den E-Mails ein knappes, sachliches Morgenbriefing auf Deutsch. "
        f"Heutiges Datum: {heute}. Loese relative Angaben (heute, morgen, Freitag, "
        "naechste Woche) anhand des heutigen Datums auf.\n"
        "Regeln:\n"
        "- Ordne jeden Punkt einem Projekt aus der Liste zu (nur den Namen); passt "
        "keins, schreibe 'Allgemein'.\n"
        "- dringlichkeit: 'Hoch' (Maengel, Behinderung, Fristen, Eskalation, Termin "
        "heute/morgen), 'Mittel' (Rechnung, Lieferavis, Antwort noetig), 'Niedrig' (Info).\n"
        "- termine NUR, wenn ein konkretes Datum oder eine Frist genannt ist. datum als "
        "YYYY-MM-DD. uhrzeit 'HH:MM' oder leer (dann ganztaegig). dauer_min Standard 60.\n"
        "- Fasse zusammen, erfinde nichts. Antworte AUSSCHLIESSLICH als JSON nach diesem Schema:\n"
        '{"ueberblick": "2-4 Saetze Gesamtlage", '
        '"punkte": [{"projekt": "", "dringlichkeit": "Hoch|Mittel|Niedrig", '
        '"richtung": "Eingang|Gesendet", "thema": "", "naechster_schritt": ""}], '
        '"termine": [{"titel": "", "datum": "YYYY-MM-DD", "uhrzeit": "", '
        '"dauer_min": 60, "ort": "", "quelle": ""}]}'
    )
    zeilen = []
    if projektnamen:
        zeilen.append("Bekannte Projekte:\n" + "\n".join(projektnamen) + "\n")
    zeilen.append(f"E-Mails ({len(mails)}):")
    for i, m in enumerate(mails, 1):
        zeilen.append(
            f"[{i}] ({m['richtung']}) {m['zeit']:%Y-%m-%d %H:%M} | {m['partner']} | "
            f"Betreff: {m['betreff']}\nAuszug: {m['auszug']}")
    daten = gemini_json(api_key, system, "\n".join(zeilen), cfg["gemini_modell"])
    daten.setdefault("ueberblick", "")
    daten.setdefault("punkte", [])
    daten.setdefault("termine", [])
    return daten


# ---------------------------------------------------------------------------
# HTML-Briefing (KPC-Design)
# ---------------------------------------------------------------------------
def baue_html(brief, mails, zeitraum):
    def esc(x):
        return html.escape(str(x or ""))

    nach_gruppe = {g: [] for g in GRUPPEN}
    for p in brief["punkte"]:
        nach_gruppe.get(p.get("dringlichkeit", "Niedrig"), nach_gruppe["Niedrig"]).append(p)
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
  h2.termine { background:#3c5a4a; }
  table { width:100%; border-collapse:collapse; font-size:12px; }
  th { text-align:left; background:#6b5040; color:#fff; font-weight:400; padding:6px 8px; }
  td { padding:6px 8px; border-bottom:1px solid #e6ddcf; vertical-align:top; }
  tr:nth-child(even) td { background:#faf6ef; }
  .proj { color:#6b5040; font-weight:700; }
  footer { margin-top:22px; color:#6b5040; font-size:11px;
           border-top:1px solid #c8a882; padding-top:8px; }
</style></head><body>
<header><h1>KPC &middot; Morgenbriefing</h1>
<div class="meta">Stand: """ + esc(stand) + " &nbsp;|&nbsp; Zeitraum: " + esc(zeitraum)
        + " &nbsp;|&nbsp; Mails: " + str(len(mails)) + """</div></header>"""]

    if brief.get("ueberblick"):
        teile.append('<div class="ueberblick">' + esc(brief["ueberblick"]) + "</div>")

    klasse = {"Hoch": "", "Mittel": "mittel", "Niedrig": "niedrig"}
    titel = {"Hoch": "Wichtig / dringend", "Mittel": "Zu erledigen", "Niedrig": "Information"}
    for g in GRUPPEN:
        rows = nach_gruppe[g]
        if not rows:
            continue
        teile.append(f'<h2 class="{klasse[g]}">{esc(titel[g])} ({len(rows)})</h2>')
        teile.append("<table><tr><th>Projekt</th><th>Richtung</th><th>Thema</th>"
                     "<th>Naechster Schritt</th></tr>")
        for p in rows:
            teile.append(
                "<tr>"
                f'<td><span class="proj">{esc(p.get("projekt", "Allgemein"))}</span></td>'
                f"<td>{esc(p.get('richtung', ''))}</td>"
                f"<td>{esc(p.get('thema', ''))}</td>"
                f"<td>{esc(p.get('naechster_schritt', ''))}</td>"
                "</tr>")
        teile.append("</table>")

    termine = brief.get("termine", [])
    if termine:
        teile.append(f'<h2 class="termine">Termine &amp; Fristen ({len(termine)})</h2>')
        teile.append("<table><tr><th>Datum</th><th>Uhrzeit</th><th>Titel</th>"
                     "<th>Ort</th><th>Quelle</th></tr>")
        for t in termine:
            teile.append(
                "<tr>"
                f"<td>{esc(t.get('datum', ''))}</td>"
                f"<td>{esc(t.get('uhrzeit', '') or 'ganztags')}</td>"
                f"<td>{esc(t.get('titel', ''))}</td>"
                f"<td>{esc(t.get('ort', ''))}</td>"
                f"<td>{esc(t.get('quelle', ''))}</td>"
                "</tr>")
        teile.append("</table>")

    teile.append('<footer>Erstellt aus Outlook (Eingang + Gesendet), Zusammenfassung '
                 'durch Gemini. E-Mails wurden nur gelesen; Termine nur nach '
                 'Bestaetigung im Kalender. Dies ersetzt keine eigene Pruefung.</footer>'
                 "</body></html>")
    return "".join(teile)


def schreibe_html(brief, mails, zeitraum, cfg):
    ordner = os.path.join(cfg["base_dir"], cfg["briefing_unterordner"])
    os.makedirs(ordner, exist_ok=True)
    pfad = os.path.join(ordner, f"Briefing_{dt.datetime.now():%Y%m%d_%H%M}.html")
    with open(pfad, "w", encoding="utf-8") as f:
        f.write(baue_html(brief, mails, zeitraum))
    return pfad


# ---------------------------------------------------------------------------
# Laufendes Protokoll ("Gedaechtnis")
# ---------------------------------------------------------------------------
def schreibe_log(brief, mails, cfg):
    pfad = os.path.join(HIER, "briefing_log.md")
    with open(pfad, "a", encoding="utf-8") as f:
        f.write(f"\n\n## {dt.datetime.now():%Y-%m-%d %H:%M}  ({len(mails)} Mails)\n\n")
        if brief.get("ueberblick"):
            f.write(brief["ueberblick"] + "\n\n")
        for p in brief.get("punkte", []):
            f.write(f"- [{p.get('dringlichkeit', '')}] {p.get('projekt', '')}: "
                    f"{p.get('thema', '')} -> {p.get('naechster_schritt', '')}\n")
        for t in brief.get("termine", []):
            f.write(f"- TERMIN {t.get('datum', '')} {t.get('uhrzeit', '')} "
                    f"{t.get('titel', '')} ({t.get('quelle', '')})\n")
    return pfad


# ---------------------------------------------------------------------------
# Termine bestaetigen + in Outlook-Kalender eintragen
# ---------------------------------------------------------------------------
def bestaetige_termine(termine):
    if not termine:
        return []
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:  # noqa: BLE001
        print("Tkinter fehlt - Termine werden nicht eingetragen.")
        return []
    root = tk.Tk()
    root.title("Termine in den Outlook-Kalender uebernehmen")
    root.geometry("780x440")
    ttk.Label(root, padding=8, text=("Haken setzen bei den Terminen, die in deinen "
              "Outlook-Kalender sollen:")).pack(anchor="w")
    rahmen = ttk.Frame(root, padding=(10, 0))
    rahmen.pack(fill="both", expand=True)
    eintraege = []
    for t in termine:
        v = tk.BooleanVar(value=True)
        txt = (f"{t.get('datum', '')}  {t.get('uhrzeit', '') or '(ganztags)'}  -  "
               f"{t.get('titel', '')}   [{(t.get('quelle', '') or '')[:45]}]")
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


def trage_termine_ein(termine):
    if not termine:
        return 0
    import pywintypes
    app = outlook_app()
    n = 0
    for t in termine:
        try:
            start, hat_uhrzeit = _start_zeit(t["datum"], t.get("uhrzeit", ""))
            appt = app.CreateItem(1)   # 1 = olAppointmentItem
            appt.Subject = t.get("titel", "Termin")
            appt.Start = pywintypes.Time(start)
            if hat_uhrzeit:
                appt.Duration = int(t.get("dauer_min") or 60)
            else:
                appt.AllDayEvent = True
            if t.get("ort"):
                appt.Location = t["ort"]
            appt.Body = "Aus KPC-Morgenbriefing.\nQuelle: " + (t.get("quelle", "") or "")
            appt.ReminderSet = True
            appt.Save()
            n += 1
        except Exception as e:  # noqa: BLE001
            print(f"  Termin '{t.get('titel', '')}' nicht eingetragen: {e}")
    return n


# ---------------------------------------------------------------------------
# Hauptlauf
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="KPC Morgenbriefing (Outlook, read-only).")
    ap.add_argument("--seit", default=None, help="Startdatum YYYY-MM-DD.")
    ap.add_argument("--stunden", type=int, default=0, help="Rueckblick in Stunden.")
    ap.add_argument("--kein-kalender", action="store_true", help="Kein Termin-Fenster.")
    ap.add_argument("--max", type=int, default=0, help="Maximale Mailanzahl (Debug).")
    args = ap.parse_args()

    cfg = lade_config()
    key = gemini_key()
    if not key:
        print("FEHLER: Kein Gemini-Schluessel. Bitte in der Datei .env eintragen:\n"
              "  GEMINI_API_KEY=AIza...\n  (Schluessel: https://aistudio.google.com/apikey)")
        sys.exit(1)

    state = lade_state()
    jetzt = dt.datetime.now()
    if args.seit:
        marker = dt.datetime.strptime(args.seit, "%Y-%m-%d")
    elif args.stunden:
        marker = jetzt - dt.timedelta(hours=args.stunden)
    elif state.get("letzter_lauf"):
        marker = dt.datetime.fromisoformat(state["letzter_lauf"])
    else:
        marker = jetzt - dt.timedelta(hours=cfg["stunden_rueckblick_erststart"])

    print(f"Lese Outlook ab {marker:%d.%m.%Y %H:%M} ...")
    mails = sammle_mails(cfg, marker, args.max or cfg["max_mails"])
    if not mails:
        print("Keine neuen Mails seit dem letzten Briefing.")
        return

    print(f"{len(mails)} Mail(s) gefunden. Erstelle Briefing mit Gemini ...")
    brief = baue_briefing(mails, lade_projektnamen(cfg), cfg, key)

    zeitraum = f"{marker:%d.%m.%Y %H:%M} - {jetzt:%d.%m.%Y %H:%M}"
    html_pfad = schreibe_html(brief, mails, zeitraum, cfg)
    schreibe_log(brief, mails, cfg)
    print(f"Briefing: {html_pfad}")
    try:
        os.startfile(html_pfad)   # nur Windows
    except Exception:  # noqa: BLE001
        pass

    state["letzter_lauf"] = jetzt.isoformat()
    speichere_state(state)

    if not args.kein_kalender:
        auswahl = bestaetige_termine(brief.get("termine", []))
        n = trage_termine_ein(auswahl)
        if n:
            print(f"{n} Termin(e) in den Outlook-Kalender eingetragen.")
        else:
            print("Keine Termine eingetragen.")


if __name__ == "__main__":
    main()
