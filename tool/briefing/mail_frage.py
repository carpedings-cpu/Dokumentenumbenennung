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


def _tokens(frage):
    """Suchbegriffe aus der Frage ziehen (Umlaute vereinfacht, Fuellwoerter raus)."""
    t = frage.lower()
    t = (t.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
         .replace("ß", "ss"))
    woerter = re.findall(r"[a-z0-9][a-z0-9._-]{2,}", t)
    return [w for w in woerter if w not in _STOPP]


def _vereinfacht(s):
    s = (s or "").lower()
    return (s.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
            .replace("ß", "ss"))


def suche_mails(frage, tage, status=None):
    """Durchsucht Posteingang + Gesendete nach zur Frage passenden Mails."""
    tokens = _tokens(frage)
    if not tokens:
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
                punkte = sum(1 for tok in tokens if tok in heuhaufen)
                if punkte == 0:
                    item = items.GetNext(); continue
                treffer.append({
                    "punkte": punkte, "zeit": zeit, "richtung": richtung,
                    "partner": partner, "betreff": betreff,
                    "auszug": re.sub(r"[ \t]+", " ", body).strip()[:900],
                })
            except Exception:  # noqa: BLE001
                pass
            item = items.GetNext()
    treffer.sort(key=lambda m: (m["punkte"], m["zeit"]), reverse=True)
    treffer = treffer[:MAX_TREFFER]
    treffer.sort(key=lambda m: m["zeit"])   # chronologisch fuer die KI
    return treffer


def gemini_antwort(api_key, frage, mails):
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


def frage_beantworten(frage, tage, status=None):
    """Kompletter Ablauf: suchen -> KI fragen -> Antwort + Quellen als Text."""
    key = B.gemini_key()
    if key in ("", "AIza...") or len(key) < 20:
        raise RuntimeError("API_KEY_INVALID: kein gueltiger Schluessel in der .env "
                           f"(gesucht in {B.HIER}).")
    if status:
        status("Durchsuche Outlook ...")
    mails = suche_mails(frage, tage, status=status)
    if not mails:
        return ("Dazu habe ich in den letzten "
                f"{tage} Tagen keine passenden E-Mails gefunden.\n\n"
                "Tipp: Anderen Suchbegriff versuchen (Projektname, Firma, "
                "Kommissionsnummer) oder den Zeitraum vergroessern.")
    if status:
        status(f"{len(mails)} passende Mails gefunden - frage die KI ...")
    antwort = gemini_antwort(B.gemini_key(), frage, mails)
    quellen = "\n".join(
        f"  - {m['zeit']:%d.%m.%Y} | {m['richtung']} | {m['partner']} | {m['betreff']}"
        for m in sorted(mails, key=lambda x: x["zeit"], reverse=True)[:12])
    return antwort + "\n\n" + "-" * 60 + f"\nGefundene Mails ({len(mails)}):\n" + quellen


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
    ttk.Label(kopf, text="Frage (z. B. 'Wie ist der letzte Stand bei Projekt Goetze?'):",
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
    status_var = tk.StringVar(value="Bereit. Outlook muss geoeffnet sein.")
    ttk.Label(root, textvariable=status_var, padding=(10, 0)).pack(anchor="w")

    ausgabe = ScrolledText(root, wrap="word", font=("", 10), state="disabled")
    ausgabe.pack(fill="both", expand=True, padx=10, pady=8)

    lauf = {"aktiv": False}

    def zeige(text):
        ausgabe.configure(state="normal")
        ausgabe.delete("1.0", "end")
        ausgabe.insert("1.0", text)
        ausgabe.configure(state="disabled")

    def setze_status(text):
        root.after(0, lambda: status_var.set(text))

    def fertig(text, fehler=False):
        def _f():
            lauf["aktiv"] = False
            knopf.configure(state="normal")
            status_var.set("Fehler - siehe unten." if fehler else "Fertig.")
            zeige(text)
        root.after(0, _f)

    def arbeiter(frage, tage):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:  # noqa: BLE001
            pass
        try:
            fertig(frage_beantworten(frage, tage, status=setze_status))
        except Exception as e:  # noqa: BLE001
            titel, text, _ = B._erklaere_fehler(e)
            fertig(f"{titel}\n\n{text}", fehler=True)

    def start(*_):
        if lauf["aktiv"]:
            return
        frage = frage_var.get().strip()
        if len(frage) < 4:
            zeige("Bitte zuerst eine Frage eintippen.")
            return
        tage = int(zeitraum.get().split()[0])
        lauf["aktiv"] = True
        knopf.configure(state="disabled")
        zeige("Einen Moment - ich durchsuche das Postfach und frage die KI ...")
        threading.Thread(target=arbeiter, args=(frage, tage), daemon=True).start()

    knopf.configure(command=start)
    eingabe.bind("<Return>", start)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        B._melde(*B._erklaere_fehler(e))
        sys.exit(1)
