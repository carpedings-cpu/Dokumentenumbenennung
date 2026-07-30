#!/usr/bin/env python3
"""
KPC Mail-Frage - Fragen an das eigene Postfach stellen.

Kleines Fenster: Frage eintippen (z. B. "Wie ist der letzte Stand bei
Projekt Goetze?"), Zeitraum waehlen, Antwort holen. Das Programm durchsucht
Outlook (Posteingang + Gesendete) LIVE nach passenden Mails, gibt die
Treffer an Gemini und zeigt die Antwort mit Quellenangaben (Datum, Absender,
Betreff). Outlook wird dabei nur GELESEN.

Benoetigt dieselbe .env wie das Morgenbriefing (GEMINI_API_KEY) - am besten
liegt die .exe im selben Ordner wie KPC-Morgenbriefing.exe.
"""

import datetime as dt
import json
import re
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request

import briefing as B

MODELL = "gemini-2.5-flash"
MAX_PRUEFEN_JE_ORDNER = 800     # so viele juengste Mails werden hoechstens angesehen
MAX_TREFFER = 40                # so viele Mails gehen hoechstens an die KI

# Woerter, die fuer die Suche nichts taugen (Fuellwoerter der Frage).
_STOPP = {
    "der", "die", "das", "den", "dem", "des", "ein", "eine", "einen", "einem",
    "und", "oder", "aber", "wie", "was", "wer", "wann", "wo", "warum", "wieso",
    "ist", "sind", "war", "waren", "wird", "werden", "wurde", "hat", "hatte",
    "haben", "kann", "koennen", "soll", "sollen", "muss", "bei", "beim", "zum",
    "zur", "mit", "von", "vom", "fuer", "auf", "aus", "nach", "ueber", "unter",
    "gibt", "es", "sich", "noch", "schon", "auch", "nicht", "mir", "mich",
    "uns", "ich", "du", "er", "sie", "wir", "ihr", "mein", "meine", "unser",
    "bitte", "mal", "denn", "dass", "im", "in", "am", "an", "um", "zu",
    "projekt", "stand", "letzte", "letzter", "letzten", "aktuell", "aktuelle",
    "aktueller", "neues", "neuen", "thema", "mail", "mails", "email", "emails",
    "nachricht", "nachrichten", "status",
}


# Abkuerzungen/Synonyme: wer nach dem einen fragt, findet auch das andere.
_SYNONYM = {
    "ibn": ["inbetriebnahme"], "inbetriebnahme": ["ibn"],
    "lv": ["leistungsverzeichnis"], "leistungsverzeichnis": ["lv"],
    "bzp": ["bauzeitenplan"], "bauzeitenplan": ["bzp"],
    "va": ["verfahrensanweisung"], "verfahrensanweisung": ["va"],
    "ag": ["auftraggeber"], "auftraggeber": ["ag"],
    "mangel": ["maengel", "mangelanzeige"], "maengel": ["mangel"],
    "abnahme": ["abnahmeprotokoll"], "rechnung": ["re.", "rechnungen"],
    "ffm": ["frankfurt"], "frankfurt": ["ffm"],
}


def _tokens(frage):
    """Suchbegriff-Gruppen aus der Frage ziehen.

    Jede Gruppe = Begriff + seine Abkuerzungen/Synonyme; eine Mail passt zur
    Gruppe, wenn IRGENDEINE Variante (auch als Teilwort) vorkommt.
    Umlaute werden vereinfacht (ae/oe/ue), Fuellwoerter fliegen raus.
    """
    t = frage.lower()
    t = (t.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
         .replace("ß", "ss"))
    woerter = re.findall(r"[a-z0-9][a-z0-9._-]+", t)
    gruppen = []
    for w in woerter:
        if w in _STOPP:
            continue
        gruppen.append({w, *_SYNONYM.get(w, [])})
    return gruppen


def _vereinfacht(s):
    s = (s or "").lower()
    return (s.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
            .replace("ß", "ss"))


def suche_mails(frage, tage, status=None):
    """Durchsucht Posteingang + Gesendete nach zur Frage passenden Mails."""
    gruppen = _tokens(frage)
    if not gruppen:
        raise RuntimeError("Bitte einen konkreten Begriff in die Frage aufnehmen "
                           "(z. B. Projektname, Firma, Ort oder Stichwort).")
    grenze = dt.datetime.now() - dt.timedelta(days=tage)
    app = B.outlook_app()
    treffer = []
    for fid, richtung in ((6, "Eingang"), (5, "Gesendet")):
        try:
            ordner = B._ordner(app, fid)
        except Exception:  # noqa: BLE001
            continue
        items = ordner.Items
        try:
            items.Sort("[SentOn]" if richtung == "Gesendet" else "[ReceivedTime]", True)
        except Exception:  # noqa: BLE001
            items.Sort("[ReceivedTime]", True)
        item = items.GetFirst()
        geprueft = 0
        while item is not None and geprueft < MAX_PRUEFEN_JE_ORDNER:
            try:
                if int(getattr(item, "Class", 0)) != 43:
                    item = items.GetNext(); continue
                zeit = B.py_datetime(getattr(item, "SentOn", None) or item.ReceivedTime)
                if zeit < grenze:
                    break
                geprueft += 1
                if status and geprueft % 100 == 0:
                    status(f"Durchsuche {richtung}: {geprueft} Mails geprueft, "
                           f"{len(treffer)} Treffer ...")
                betreff = item.Subject or ""
                try:
                    body = (item.Body or "")[:4000]
                except Exception:  # noqa: BLE001
                    body = ""
                if richtung == "Gesendet":
                    partner = "An: " + B.empfaenger_anzeige(item)
                else:
                    partner = B.smtp_adresse(item) or "(unbekannt)"
                heuhaufen = _vereinfacht(betreff + " " + partner + " " + body)
                punkte = sum(1 for g in gruppen
                             if any(tok in heuhaufen for tok in g))
                if punkte == 0:
                    item = items.GetNext(); continue
                try:
                    entry_id = item.EntryID or ""
                except Exception:  # noqa: BLE001
                    entry_id = ""
                try:
                    store_id = item.Parent.StoreID or ""
                except Exception:  # noqa: BLE001
                    store_id = ""
                treffer.append({
                    "punkte": punkte, "zeit": zeit, "richtung": richtung,
                    "partner": partner, "betreff": betreff,
                    "auszug": re.sub(r"[ \t]+", " ", body).strip()[:900],
                    "entry_id": entry_id, "store_id": store_id,
                })
            except Exception:  # noqa: BLE001
                pass
            item = items.GetNext()
    treffer.sort(key=lambda m: (m["punkte"], m["zeit"]), reverse=True)
    treffer = treffer[:MAX_TREFFER]
    treffer.sort(key=lambda m: m["zeit"])   # chronologisch fuer die KI
    return treffer


def gemini_antwort(api_key, frage, mails, system_text=None):
    if system_text is None:
        system_text = (
            "Du bist ein Assistent fuer einen Projektleiter im Kuechen-/Anlagenbau. "
            "Beantworte die Frage AUSSCHLIESSLICH anhand der mitgelieferten "
            "E-Mail-Auszuege. Antworte auf Deutsch, kurz und klar. Beginne mit dem "
            "AKTUELLEN Stand (neueste Information), danach ggf. kurz die Vorgeschichte. "
            "Nenne zu jeder Kernaussage Datum und Absender in Klammern, z. B. "
            "(04.07., meier@bau.de). Wenn die Auszuege die Frage nicht beantworten, "
            "sage das ehrlich und rate nicht."
        )
    zeilen = []
    for i, m in enumerate(mails, 1):
        zeilen.append(f"[{i}] {m['zeit']:%d.%m.%Y} | {m['richtung']} | "
                      f"{m['partner']} | Betreff: {m['betreff']}\n{m['auszug']}")
    user_text = f"FRAGE: {frage}\n\nE-MAILS ({len(mails)}):\n\n" + "\n\n".join(zeilen)
    body = {
        "contents": [{"parts": [{"text": user_text}]}],
        "system_instruction": {"parts": [{"text": system_text}]},
        "generationConfig": {"temperature": 0.3},
    }
    url = B.GEMINI_ENDPUNKT.format(m=MODELL) + "?key=" + urllib.parse.quote(api_key)
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"},
                                 method="POST")
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
        return "".join(p.get("text", "") for p in cand["content"]["parts"]).strip()
    except (KeyError, IndexError) as e:
        raise RuntimeError(f"Unerwartete Gemini-Antwort: {json.dumps(daten)[:300]}") from e


_BRIEFING_SYSTEM = (
    "Du bist ein Assistent fuer einen Projektleiter im Kuechen-/Anlagenbau. "
    "Erstelle ein KOMPAKTES PROJEKT-BRIEFING AUSSCHLIESSLICH aus den "
    "mitgelieferten E-Mail-Auszuegen zu dem genannten Projekt. Deutsch, knapp, "
    "keine Floskeln. Gliederung genau so:\n\n"
    "AKTUELLER STAND\n3-5 Saetze, das Neueste zuerst.\n\n"
    "OFFENE PUNKTE / ZU TUN\nStichpunkte: was ist zu erledigen, wer ist dran, "
    "je mit (Datum, Absender). Nur wirklich Offenes - Erledigtes weglassen.\n\n"
    "WARTET AUF ANTWORT\nWer schuldet wem seit wann eine Antwort.\n\n"
    "Wenn die Auszuege zu einem Abschnitt nichts hergeben, schreibe dort "
    "'Nichts Offenes gefunden.' und erfinde nichts."
)


def frage_beantworten(frage, tage, status=None, modus="frage"):
    """Kompletter Ablauf: suchen -> KI fragen. Gibt (Antworttext, Mails) zurueck.

    modus 'frage'    -> freie Frage beantworten
    modus 'briefing' -> strukturiertes Projekt-Briefing (Stand/To-do/Wartet)
    """
    key = B.gemini_key()
    if key in ("", "AIza...") or len(key) < 20:
        raise RuntimeError("API_KEY_INVALID: kein gueltiger Schluessel in der .env "
                           f"(gesucht in {B.HIER}).")
    if status:
        status("Durchsuche Outlook ...")
    mails = suche_mails(frage, tage, status=status)
    if not mails:
        return (("Dazu habe ich in den letzten "
                 f"{tage} Tagen keine passenden E-Mails gefunden.\n\n"
                 "Tipp: Anderen Suchbegriff versuchen (Projektname, Firma, "
                 "Kommissionsnummer) oder den Zeitraum vergroessern."), [])
    if status:
        status(f"{len(mails)} passende Mails gefunden - frage die KI ...")
    if modus == "briefing":
        antwort = gemini_antwort(B.gemini_key(),
                                 f"Projekt-Briefing fuer: {frage}",
                                 mails, system_text=_BRIEFING_SYSTEM)
    else:
        antwort = gemini_antwort(B.gemini_key(), frage, mails)
    return antwort, sorted(mails, key=lambda x: x["zeit"], reverse=True)


# ---------------------------------------------------------------------------
# Fenster
# ---------------------------------------------------------------------------
def main():
    import tkinter as tk
    from tkinter import ttk
    from tkinter.scrolledtext import ScrolledText

    root = tk.Tk()
    root.title("KPC Mail-Frage - Fragen an das Postfach")
    root.geometry("860x620")
    root.minsize(700, 480)

    kopf = ttk.Frame(root, padding=10)
    kopf.pack(fill="x")
    ttk.Label(kopf, text=("Frage stellen ODER Projektname eingeben und "
                          "'Projekt-Briefing' klicken (z. B. 'Jüdische Akademie FfM'):"),
              font=("", 10, "bold")).pack(anchor="w")
    frage_var = tk.StringVar()
    eingabe = ttk.Entry(kopf, textvariable=frage_var, font=("", 11))
    eingabe.pack(fill="x", pady=(4, 6))
    eingabe.focus_set()

    zeile = ttk.Frame(kopf)
    zeile.pack(fill="x")
    ttk.Label(zeile, text="Zeitraum:").pack(side="left")
    zeitraum = ttk.Combobox(zeile, state="readonly", width=16,
                            values=["30 Tage", "90 Tage", "180 Tage", "365 Tage"])
    zeitraum.set("90 Tage")
    zeitraum.pack(side="left", padx=6)
    knopf = ttk.Button(zeile, text="Antwort holen")
    knopf.pack(side="left", padx=10)
    knopf_brief = ttk.Button(zeile, text="Projekt-Briefing")
    knopf_brief.pack(side="left")
    status_var = tk.StringVar(value="Bereit. Outlook muss geoeffnet sein.")
    ttk.Label(root, textvariable=status_var, padding=(10, 0)).pack(anchor="w")

    ausgabe = ScrolledText(root, wrap="word", font=("", 10), state="disabled")
    ausgabe.pack(fill="both", expand=True, padx=10, pady=(8, 4))

    ttk.Label(root, padding=(10, 0), font=("", 9, "bold"),
              text="Gefundene Mails - anklicken öffnet die Mail in Outlook:"
              ).pack(anchor="w")
    quell_rahmen = ttk.Frame(root, padding=(10, 2, 10, 8))
    quell_rahmen.pack(fill="x")
    quellen = ttk.Treeview(quell_rahmen, columns=("datum", "richtung", "partner", "betreff"),
                           show="headings", height=7)
    for c, t, w in (("datum", "Datum", 90), ("richtung", "Richtung", 80),
                    ("partner", "Von / An", 220), ("betreff", "Betreff", 400)):
        quellen.heading(c, text=t)
        quellen.column(c, width=w, anchor="w")
    qsb = ttk.Scrollbar(quell_rahmen, orient="vertical", command=quellen.yview)
    quellen.configure(yscrollcommand=qsb.set)
    quellen.pack(side="left", fill="x", expand=True)
    qsb.pack(side="right", fill="y")

    lauf = {"aktiv": False}
    mail_von_zeile = {}

    def zeige(text):
        ausgabe.configure(state="normal")
        ausgabe.delete("1.0", "end")
        ausgabe.insert("1.0", text)
        ausgabe.configure(state="disabled")

    def fuelle_quellen(mails):
        mail_von_zeile.clear()
        for iid in quellen.get_children():
            quellen.delete(iid)
        for m in mails:
            iid = quellen.insert("", "end", values=(
                f"{m['zeit']:%d.%m.%Y}", m["richtung"], m["partner"], m["betreff"]))
            mail_von_zeile[iid] = m

    def oeffne_quelle(event):
        from tkinter import messagebox
        iid = quellen.identify_row(event.y)
        if not iid:
            return                      # Klick auf Kopfzeile/leeren Bereich
        m = mail_von_zeile.get(iid)
        if not m:
            messagebox.showinfo("Hinweis", "Bitte zuerst eine Frage stellen - "
                                "dann erscheinen hier anklickbare Mails.")
            return
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:  # noqa: BLE001
            pass
        if not B.oeffne_mail_in_outlook(m.get("entry_id", ""), m.get("store_id", "")):
            messagebox.showerror(
                "Mail nicht gefunden",
                "Die Mail konnte nicht geöffnet werden - ist Outlook noch "
                "geöffnet? Eventuell wurde sie verschoben oder gelöscht.\n\n"
                f"Technische Meldung: {B.LETZTER_OEFFNEN_FEHLER or '-'}")
    # Einfacher Klick genuegt (Wunsch): beim Loslassen der Maustaste oeffnen.
    quellen.bind("<ButtonRelease-1>", oeffne_quelle)

    def setze_status(text):
        root.after(0, lambda: status_var.set(text))

    def fertig(text, mails=None, fehler=False):
        def _f():
            lauf["aktiv"] = False
            knopf.configure(state="normal")
            knopf_brief.configure(state="normal")
            status_var.set("Fehler - siehe unten." if fehler else "Fertig.")
            zeige(text)
            fuelle_quellen(mails or [])
        root.after(0, _f)

    def arbeiter(frage, tage, modus):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:  # noqa: BLE001
            pass
        try:
            antwort, mails = frage_beantworten(frage, tage, status=setze_status,
                                               modus=modus)
            fertig(antwort, mails)
        except Exception as e:  # noqa: BLE001
            titel, text, _ = B._erklaere_fehler(e)
            fertig(f"{titel}\n\n{text}", fehler=True)

    def start(modus="frage"):
        if lauf["aktiv"]:
            return
        frage = frage_var.get().strip()
        if len(frage) < 4:
            zeige("Bitte zuerst eine Frage bzw. einen Projektnamen eintippen.")
            return
        tage = int(zeitraum.get().split()[0])
        lauf["aktiv"] = True
        knopf.configure(state="disabled")
        knopf_brief.configure(state="disabled")
        zeige("Einen Moment - ich durchsuche das Postfach und "
              + ("erstelle das Projekt-Briefing ..." if modus == "briefing"
                 else "frage die KI ..."))
        threading.Thread(target=arbeiter, args=(frage, tage, modus),
                         daemon=True).start()

    knopf.configure(command=lambda: start("frage"))
    knopf_brief.configure(command=lambda: start("briefing"))
    eingabe.bind("<Return>", lambda *_: start("frage"))
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        B._melde(*B._erklaere_fehler(e))
        sys.exit(1)
