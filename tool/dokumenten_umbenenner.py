#!/usr/bin/env python3
"""
Dokumenten-Umbenenner (KPC VA 1.1) - eigenstaendiges Desktop-Programm.

Benennt Projektdokumente nach der Verfahrensanweisung "Dokumentenbenennung"
um. Zwei Modi (Umschalter):
  - Offline (regelbasiert): liest Datum automatisch, schlaegt einen Typ vor;
    Felder werden bestaetigt/ergaenzt. Keine Installation, kein Internet.
  - Claude-API: erkennt Datum/Typ/Bezeichnung automatisch (auch bei Scans).
    Benoetigt einen API-Schluessel und das Paket 'anthropic'.

Start:  python dokumenten_umbenenner.py
"""

import os
import re
import shutil
import sys

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ModuleNotFoundError:  # Tkinter fehlt (z. B. minimales Linux ohne python3-tk)
    tk = None

# Optionales Drag & Drop (Dateien/Ordner ins Fenster ziehen).
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:
    TkinterDnD = None
    DND_FILES = None
    _HAS_DND = False

# Skriptverzeichnis in den Pfad, damit die Module unabhaengig vom CWD laden.
_HIER = os.path.dirname(os.path.abspath(__file__))
if _HIER not in sys.path:
    sys.path.insert(0, _HIER)

import va_rules as VA            # noqa: E402
import email_extract             # noqa: E402
import projekt_zuordnung as PZ   # noqa: E402
from pdf_text import extract_pdf_text  # noqa: E402

LESBARE_TEXT_ENDUNGEN = {".pdf", ".txt", ".md"}
ANALYSIERBAR = LESBARE_TEXT_ENDUNGEN | {".png", ".jpg", ".jpeg", ".gif", ".webp"}

# Standard-Projektbasis (enthaelt die Projekt-Unterordner). Identisch zur Triage;
# im Programm aenderbar und wird automatisch korrigiert, wenn ein Ordner namens
# "00_Posteingang" eingelesen wird (dann = dessen uebergeordneter Ordner).
STANDARD_BASIS = r"C:\Users\ziegler\Desktop\Dokumentenumbenennung"
EINGANG_ORDNER = "00_Posteingang"


def _betreff_kurz(betreff):
    """E-Mail-Betreff als Bezeichnung: Antwort-/Weiterleitungs-Präfixe weg, gekürzt."""
    s = (betreff or "").strip()
    while True:
        neu = re.sub(r"(?i)^\s*(aw|wg|re|fw|fwd|antw|antwort)\s*:\s*", "", s)
        if neu == s:
            break
        s = neu
    return s[:60].strip()


def lade_va_regeln():
    """Laedt den VA-Text aus der SKILL.md (Single Source of Truth) als Systemprompt."""
    kandidaten = []
    # In einer PyInstaller-.exe liegt die SKILL.md gebuendelt in sys._MEIPASS.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        kandidaten.append(os.path.join(meipass, "SKILL.md"))
    kandidaten.append(os.path.join(_HIER, "..", ".claude", "skills",
                                   "dokumentenumbenennung", "SKILL.md"))
    for pfad in kandidaten:
        try:
            with open(pfad, "r", encoding="utf-8") as f:
                txt = f.read()
            if txt.startswith("---"):
                teile = txt.split("---", 2)
                if len(teile) == 3:
                    txt = teile[2]
            return txt.strip()
        except OSError:
            continue
    return ("Verfahrensanweisung KPC Dokumentenbenennung 1.1. Schema: "
            "JJMMTT_[Quelle]_[Phase]_Dokumententyp_Bezeichnung[_Version].")


class App:
    def __init__(self, root):
        self.root = root
        root.title("Dokumenten-Umbenenner – KPC VA 1.1")
        root.geometry("1080x720")

        self.ordner = tk.StringVar()
        self.modus = tk.StringVar(value="offline")
        self.api_key = tk.StringVar(value=os.environ.get("ANTHROPIC_API_KEY", ""))
        self.einsortieren = tk.BooleanVar(value=True)   # in Projektordner einsortieren
        self.projektbasis = tk.StringVar(value=STANDARD_BASIS)
        self.rows = {}  # tree-item-id -> dict mit Feldern
        self._pz_base = None    # Cache: zuletzt geladene Projektbasis
        self._pz_cache = None   # Cache: Projektliste

        self._baue_oben()
        self._baue_tabelle()
        self._baue_details()
        self._baue_unten()
        self._aktiviere_dnd()

    def _aktiviere_dnd(self):
        """Registriert das Fenster fuer Drag & Drop; meldet Erfolg/Misserfolg."""
        self.dnd_ok = False
        if not _HAS_DND:
            self.protokoll("Drag & Drop nicht verfügbar – bitte 'Dateien wählen…' nutzen.")
            return
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop)
            self.dnd_ok = True
            self.protokoll("Drag & Drop aktiv. Hinweis: Outlook-Mails ggf. erst "
                           "als Datei (.msg) speichern, dann hineinziehen.")
        except Exception as e:  # noqa: BLE001
            self.protokoll(f"Drag & Drop konnte nicht aktiviert werden: {e} "
                           "– bitte 'Dateien wählen…' nutzen.")

    # ----------------------------------------------------------------- UI
    def _baue_oben(self):
        f = ttk.Frame(self.root, padding=8)
        f.pack(fill="x")
        ttk.Label(f, text="Ordner:").pack(side="left")
        ttk.Entry(f, textvariable=self.ordner, width=70).pack(side="left", padx=4)
        ttk.Button(f, text="Durchsuchen…", command=self.waehle_ordner).pack(side="left")
        ttk.Button(f, text="Einlesen", command=self.einlesen).pack(side="left", padx=4)
        ttk.Button(f, text="Dateien wählen…", command=self.waehle_dateien).pack(side="left")
        ttk.Button(f, text="Liste leeren", command=self._leeren).pack(side="left", padx=4)

        m = ttk.Frame(self.root, padding=(8, 0))
        m.pack(fill="x")
        ttk.Label(m, text="Modus:").pack(side="left")
        ttk.Radiobutton(m, text="Offline (regelbasiert)", value="offline",
                        variable=self.modus).pack(side="left")
        ttk.Radiobutton(m, text="Claude-API", value="claude",
                        variable=self.modus).pack(side="left", padx=(4, 0))
        ttk.Radiobutton(m, text="Gemini-API", value="gemini",
                        variable=self.modus).pack(side="left", padx=(4, 12))
        ttk.Label(m, text="API-Schlüssel:").pack(side="left")
        ttk.Entry(m, textvariable=self.api_key, width=32, show="•").pack(side="left", padx=4)

        s = ttk.Frame(self.root, padding=(8, 2))
        s.pack(fill="x")
        ttk.Checkbutton(s, text="Umbenannte Dateien in Projektordner einsortieren",
                        variable=self.einsortieren).pack(side="left")
        ttk.Label(s, text="Projektbasis:").pack(side="left", padx=(12, 2))
        ttk.Entry(s, textvariable=self.projektbasis, width=48).pack(side="left", padx=2)
        ttk.Button(s, text="Durchsuchen…", command=self.waehle_basis).pack(side="left")

    def _baue_tabelle(self):
        hinweis = ("Reinziehen ODER 'Dateien wählen…' (auch .eml/.msg). E-Mails werden "
                   "in Mailtext-PDF + Anhänge zerlegt. Outlook-Mail ggf. erst als .msg "
                   "speichern.") if _HAS_DND else \
                  ("'Dateien wählen…' oder Ordner wählen + 'Einlesen' (auch .eml/.msg).")
        ttk.Label(self.root, text=hinweis).pack(anchor="w", padx=10, pady=(0, 2))

        f = ttk.Frame(self.root, padding=8)
        f.pack(fill="both", expand=True)
        cols = ("alt", "typ", "datum", "projekt", "neu")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=12)
        for c, t, w in (("alt", "Alt (Ist)", 300), ("typ", "Typ", 130),
                        ("datum", "Datum", 70), ("projekt", "Projektordner", 200),
                        ("neu", "Neu (Vorschau)", 360)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.zeige_auswahl)

    def _baue_details(self):
        f = ttk.LabelFrame(self.root, text="Felder der ausgewählten Datei", padding=8)
        f.pack(fill="x", padx=8)

        self.var = {k: tk.StringVar() for k in (
            "datum", "quelle", "phase", "dokumententyp", "bezeichnung", "version",
            "projekt",
            "projektnr", "auftragsart", "leistungsphase", "planinhalt",
            "geschoss", "plannr", "index", "status")}
        self.ist_plan = tk.BooleanVar(value=False)

        # Standardfelder
        std = ttk.Frame(f)
        std.pack(fill="x")
        self._feld(std, "Datum (JJMMTT)", "datum", 0, 0, 10)
        self._feld(std, "Quelle", "quelle", 0, 2, 16)
        self._feld(std, "Phase", "phase", 0, 4, 14)
        ttk.Label(std, text="Dokumententyp").grid(row=1, column=0, sticky="w", padx=4)
        cb = ttk.Combobox(std, textvariable=self.var["dokumententyp"],
                          values=VA.DOKUMENTTYPEN, width=28)
        cb.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        self._feld(std, "Bezeichnung", "bezeichnung", 1, 2, 24)
        self._feld(std, "Version", "version", 1, 4, 10)
        self._feld(std, "Projektordner (Ablage)", "projekt", 2, 0, 40)

        ttk.Checkbutton(f, text="Planunterlage (eigenes Schema, Kapitel 6)",
                        variable=self.ist_plan,
                        command=self.aktualisiere_vorschau).pack(anchor="w", pady=(6, 0))

        # Planfelder
        pl = ttk.Frame(f)
        pl.pack(fill="x")
        self._feld(pl, "Projektnr.", "projektnr", 0, 0, 8)
        self._feld(pl, "Auftragsart (10/20)", "auftragsart", 0, 2, 6)
        self._feld(pl, "Leistungsphase (1–9)", "leistungsphase", 0, 4, 6)
        self._feld(pl, "Planinhalt (A,E,…)", "planinhalt", 0, 6, 6)
        self._feld(pl, "Geschoss (EG,01,…)", "geschoss", 1, 0, 8)
        self._feld(pl, "Plannr. (01)", "plannr", 1, 2, 6)
        self._feld(pl, "Index (A,B,…)", "index", 1, 4, 6)
        self._feld(pl, "Status (V/PL/F)", "status", 1, 6, 6)

        ttk.Button(f, text="Vorschau aktualisieren",
                   command=self.aktualisiere_vorschau).pack(anchor="w", pady=6)

    def _feld(self, parent, label, key, r, c, width):
        ttk.Label(parent, text=label).grid(row=r, column=c, sticky="w", padx=4)
        ttk.Entry(parent, textvariable=self.var[key], width=width)\
            .grid(row=r, column=c + 1, sticky="w", padx=4, pady=2)

    def _baue_unten(self):
        f = ttk.Frame(self.root, padding=8)
        f.pack(fill="x")
        ttk.Button(f, text="Alle umbenennen", command=self.umbenennen)\
            .pack(side="left")
        self.status = tk.StringVar(value="Bereit.")
        ttk.Label(f, textvariable=self.status).pack(side="left", padx=12)

        self.log = tk.Text(self.root, height=7)
        self.log.pack(fill="both", expand=False, padx=8, pady=(0, 8))

    # -------------------------------------------------------------- Aktionen
    def protokoll(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def waehle_ordner(self):
        d = filedialog.askdirectory(title="Ordner mit Dokumenten wählen")
        if d:
            self.ordner.set(d)
            self._basis_aus_eingang(d)

    def waehle_basis(self):
        d = filedialog.askdirectory(title="Projektbasis wählen (enthält die Projektordner)")
        if d:
            self.projektbasis.set(d)

    def _basis_aus_eingang(self, ordnerpfad):
        """Setzt die Projektbasis automatisch, wenn ein '00_Posteingang' eingelesen
        wird – dann ist die Basis dessen übergeordneter Ordner. Eine bereits
        gültige Basis wird nur überschrieben, wenn sie gar nicht existiert."""
        try:
            ordnerpfad = os.path.abspath(ordnerpfad)
            if os.path.basename(ordnerpfad).lower() != EINGANG_ORDNER.lower():
                return
            eltern = os.path.dirname(ordnerpfad)
            aktuell = self.projektbasis.get().strip()
            if eltern and os.path.isdir(eltern) and not os.path.isdir(aktuell):
                self.projektbasis.set(eltern)
        except Exception:  # noqa: BLE001
            pass

    def _projekte(self):
        """Lädt (und cached) die Projektliste der aktuellen Projektbasis."""
        base = self.projektbasis.get().strip()
        if base != self._pz_base or self._pz_cache is None:
            self._pz_base = base
            self._pz_cache = PZ.lade_projekte(base, _HIER)
            if base and not os.path.isdir(base):
                self.protokoll(f"Hinweis: Projektbasis nicht gefunden: {base} "
                               "– es wird nichts einsortiert.")
        return self._pz_cache

    def _erkenne_projekt(self, name, kontext, felder):
        """Bestimmt den Projektordner für eine Datei (oder '' wenn unklar)."""
        if not self.einsortieren.get():
            return ""
        projekte = self._projekte()
        if not projekte:
            return ""
        such = " ".join([name, kontext or "", felder.get("_text", "")])
        rec = PZ.finde_projekt(projekte, such)
        return rec["ordner"] if rec else ""

    def waehle_dateien(self):
        """Sicherer Weg ohne Drag & Drop: Dateien/E-Mails per Dialog auswählen."""
        pfade = filedialog.askopenfilenames(
            title="Dateien oder E-Mails wählen",
            filetypes=[("Alle Dateien", "*.*"),
                       ("Dokumente & E-Mails", "*.pdf *.eml *.msg *.docx *.doc *.txt "
                                              "*.png *.jpg *.jpeg *.tif *.tiff")])
        if pfade:
            self._add_files(list(pfade), anhaengen=True)

    def _dateien(self):
        d = self.ordner.get()
        if not d or not os.path.isdir(d):
            return []
        return sorted(n for n in os.listdir(d)
                      if os.path.isfile(os.path.join(d, n)) and not n.startswith("."))

    def _leeren(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.rows.clear()
        self.status.set("Liste geleert.")

    def _expandiere(self, pfade):
        """Macht aus Datei-/Ordnerpfaden eine flache Liste echter Dateien."""
        dateien = []
        for p in pfade:
            p = p.strip().strip("{}")  # tkinterdnd2 klammert Pfade mit Leerzeichen
            if os.path.isdir(p):
                for n in sorted(os.listdir(p)):
                    fp = os.path.join(p, n)
                    if os.path.isfile(fp) and not n.startswith("."):
                        dateien.append(fp)
            elif os.path.isfile(p):
                dateien.append(p)
        return dateien

    def einlesen(self):
        d = self.ordner.get()
        if not d or not os.path.isdir(d):
            messagebox.showinfo("Hinweis", "Bitte zuerst einen gültigen Ordner wählen.")
            return
        self._basis_aus_eingang(d)
        self._add_files([os.path.join(d, n) for n in self._dateien()], anhaengen=False)

    def on_drop(self, event):
        """Wird beim Hineinziehen ausgelöst. Nur echte Dateien/Ordner verarbeiten."""
        try:
            roh = list(self.root.tk.splitlist(event.data))
        except Exception:
            roh = [event.data]
        echte = [p.strip().strip("{}") for p in roh]
        echte = [p for p in echte if p and os.path.exists(p)]
        if not echte:
            self.protokoll("Drop erkannt, aber keine Datei erhalten. Wurde eine E-Mail "
                           "DIREKT aus Outlook gezogen? Bitte die Mail zuerst als .msg "
                           "speichern (Outlook: Datei → Speichern unter) und dann "
                           "hineinziehen oder 'Dateien wählen…' nutzen.")
            messagebox.showinfo(
                "Keine Datei erhalten",
                "Beim Reinziehen kam keine Datei an.\n\n"
                "Wenn du eine E-Mail direkt aus Outlook gezogen hast: Outlook gibt "
                "dabei keine Datei weiter.\n\n"
                "Speichere die Mail zuerst als .msg-Datei (Outlook: Datei -> "
                "Speichern unter, Dateityp 'Outlook-Nachricht (*.msg)') und ziehe "
                "dann diese Datei hinein - oder nutze 'Dateien wählen...'.")
            return
        self._add_files(echte, anhaengen=True)

    def _emails_extrahieren(self, dateien):
        """E-Mails (.eml/.msg) -> Mailtext-PDF + Anhaenge; sonst Datei unveraendert.

        Liefert eine Liste von (pfad, kontext, betreff_default). Der Kontext
        (Betreff + Auszug + E-Mail-Dateiname) erlaubt es, alle aus EINER Mail
        erzeugten Dateien demselben Projekt zuzuordnen; betreff_default ist nur
        für die Mailtext-PDF gesetzt (dient als Vorgabe-Bezeichnung).
        """
        ergebnis = []
        for fp in dateien:
            if email_extract.ist_email(fp):
                try:
                    neu, kontext, betreff = email_extract.extrahiere(fp)
                    ktx = (os.path.basename(fp) + " " + (kontext or "")).strip()
                    if neu:
                        self.protokoll(f"E-Mail {os.path.basename(fp)} -> {len(neu)} Datei(en) "
                                       "extrahiert (Mailtext + Anhänge; Signatur-Logos übersprungen).")
                        for n in neu:
                            bdef = betreff if "_mailtext" in os.path.basename(n).lower() else ""
                            ergebnis.append((n, ktx, bdef))
                    else:
                        self.protokoll(f"  {os.path.basename(fp)}: nichts extrahierbar.")
                except Exception as e:  # noqa: BLE001
                    self.protokoll(f"  E-Mail-Fehler bei {os.path.basename(fp)}: {e}")
            else:
                ergebnis.append((fp, "", ""))
        return ergebnis

    def _add_files(self, pfade, anhaengen=True):
        paare = self._emails_extrahieren(self._expandiere(pfade))
        if not paare:
            messagebox.showinfo("Hinweis", "Keine Dateien gefunden.")
            return
        modus = self.modus.get()
        ist_api = modus in ("claude", "gemini")
        if ist_api and not self.api_key.get().strip():
            messagebox.showwarning("API-Schlüssel fehlt",
                                   "Für die automatische Erkennung im API-Modus bitte "
                                   "oben den passenden API-Schlüssel eingeben.")
            return
        if not anhaengen:
            self._leeren()
        va_regeln = lade_va_regeln() if ist_api else None
        self.protokoll(f"Lese {len(paare)} Datei(en) – Modus: {modus} …")
        for idx, (fp, kontext, betreff_default) in enumerate(paare, 1):
            name = os.path.basename(fp)
            self.status.set(f"Analysiere {idx}/{len(paare)}: {name}")
            self.root.update_idletasks()
            felder = self._analysiere(fp, name, modus, va_regeln)
            felder["projekt"] = self._erkenne_projekt(name, kontext, felder)
            if betreff_default and not felder.get("bezeichnung"):
                felder["bezeichnung"] = _betreff_kurz(betreff_default)
            self._zeile_einfuegen(fp, felder)
        self.status.set(f"Fertig: {len(paare)} Datei(en). Felder pruefen, dann 'Alle umbenennen'.")

    def _analysiere(self, pfad, name, modus, va_regeln):
        endung = os.path.splitext(name)[1].lower()
        felder = {k: "" for k in self.var}
        felder["ist_plan"] = False
        felder["_text"] = ""    # Textauszug für die Projekt-Zuordnung

        if modus in ("claude", "gemini"):
            try:
                key = self.api_key.get().strip()
                if modus == "claude":
                    import api_client
                    erg = api_client.analysiere(pfad, va_regeln, key)
                else:
                    import gemini_client
                    erg = gemini_client.analysiere(pfad, va_regeln, key)
                for k in ("datum", "quelle", "phase", "dokumententyp",
                          "bezeichnung", "version"):
                    felder[k] = erg.get(k, "") or ""
                felder["ist_plan"] = bool(erg.get("ist_plan"))
                # --- VA-Konformitaet erzwingen ---
                roh_typ = felder["dokumententyp"]
                felder["dokumententyp"] = VA.normalisiere_typ(roh_typ)
                if roh_typ and not felder["dokumententyp"]:
                    self.protokoll(f"  {name}: Typ '{roh_typ}' ist nicht in der "
                                   "VA-Referenzliste – bitte Typ manuell wählen.")
                felder["datum"] = VA.normalisiere_datum(felder["datum"])
                if felder["dokumententyp"] in VA.TYPEN_OHNE_DATUM:
                    felder["datum"] = ""        # Datum entfaellt laut VA
                if erg.get("hinweis"):
                    self.protokoll(f"  {name}: {erg['hinweis']}")
                return felder
            except Exception as e:  # noqa: BLE001
                self.protokoll(f"  API-Fehler bei {name}: {e}")
                # weiter mit Offline-Heuristik als Rückfall

        # Offline / Rückfall
        text = ""
        if endung in (".pdf",):
            text = extract_pdf_text(pfad)
        elif endung in (".txt", ".md"):
            try:
                with open(pfad, "r", encoding="utf-8", errors="replace") as f:
                    text = f.read(20000)
            except OSError:
                text = ""
        if endung in (".pdf",) and not text:
            self.protokoll(f"  {name}: kein Text lesbar (Scan?) – Felder bitte manuell"
                           " ausfüllen oder API-Modus nutzen.")
        felder["_text"] = (text or "")[:4000]
        felder["datum"] = VA.datum_aus_text(text)
        felder["dokumententyp"] = VA.typ_aus_text(text)
        if felder["dokumententyp"] in VA.TYPEN_OHNE_DATUM:
            felder["datum"] = ""
        return felder

    def _stamm(self, felder):
        if felder.get("ist_plan"):
            return VA.baue_planname(felder)
        if not felder.get("dokumententyp"):
            return ""   # ohne gueltigen Typ kein VA-konformer Name
        return VA.baue_standardname(felder)

    def _zeile_einfuegen(self, pfad, felder):
        name = os.path.basename(pfad)
        stamm, ext = os.path.splitext(name)
        felder["_ext"] = ext
        felder["_orig"] = name
        felder["_dir"] = os.path.dirname(os.path.abspath(pfad))
        neu = self._stamm(felder)
        neu_anzeige = (neu + ext) if neu else "(unvollständig)"
        item = self.tree.insert("", "end",
                                values=(name, felder.get("dokumententyp", ""),
                                        felder.get("datum", ""),
                                        felder.get("projekt", "") or "—", neu_anzeige))
        self.rows[item] = felder

    def zeige_auswahl(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        felder = self.rows[sel[0]]
        for k in self.var:
            self.var[k].set(felder.get(k, ""))
        self.ist_plan.set(bool(felder.get("ist_plan")))

    def aktualisiere_vorschau(self):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        felder = self.rows[item]
        for k in self.var:
            felder[k] = self.var[k].get().strip()
        felder["ist_plan"] = bool(self.ist_plan.get())
        neu = self._stamm(felder)
        ext = felder["_ext"]
        anzeige = (neu + ext) if neu else "(unvollständig)"
        self.tree.item(item, values=(felder["_orig"],
                                     felder.get("dokumententyp", ""),
                                     felder.get("datum", ""),
                                     felder.get("projekt", "") or "—", anzeige))

    def _zielordner(self, felder):
        """Zielordner für eine Datei: Projektordner (falls einsortieren + erkannt
        + existiert), sonst der Quellordner."""
        if not self.einsortieren.get():
            return felder["_dir"]
        name = (felder.get("projekt") or "").strip()
        ziel = PZ.projektordner(self.projektbasis.get().strip(), name)
        return ziel or felder["_dir"]

    def umbenennen(self):
        if not self.rows:
            return
        plan = []
        belegt = {}  # Zielordner -> Set bereits geplanter Namen (Kollisionsschutz)
        for item, felder in self.rows.items():
            stamm = self._stamm(felder)
            if not stamm:
                continue
            src = felder["_dir"]
            dst = self._zielordner(felder)
            verschieben = os.path.normcase(os.path.abspath(dst)) != \
                os.path.normcase(os.path.abspath(src))
            vorhandene = belegt.setdefault(os.path.abspath(dst), set())
            # Beim Verschieben den "keine Änderung"-Vergleich abschalten.
            ref_name = "" if verschieben else felder["_orig"]
            ziel = VA.eindeutiger_zielname(dst, stamm, felder["_ext"],
                                           ref_name, vorhandene=vorhandene)
            if not verschieben and ziel == felder["_orig"]:
                continue
            plan.append((item, src, felder["_orig"], dst, ziel, verschieben))
            vorhandene.add(ziel.lower())

        if not plan:
            messagebox.showinfo("Nichts zu tun",
                                "Keine Änderungen (Felder unvollständig oder Namen gleich).")
            return

        def zeile(alt, dst, neu, verschieben):
            if verschieben:
                return f"{alt}  →  {os.path.basename(dst)}\\{neu}"
            return f"{alt}  →  {neu}"

        vorschau = "\n".join(zeile(alt, dst, neu, v)
                             for _, _, alt, dst, neu, v in plan[:25])
        if len(plan) > 25:
            vorschau += f"\n… und {len(plan) - 25} weitere"
        anzahl_verschoben = sum(1 for *_, v in plan if v)
        kopf = f"{len(plan)} Datei(en) umbenennen"
        if anzahl_verschoben:
            kopf += f" (davon {anzahl_verschoben} in Projektordner einsortieren)"
        if not messagebox.askyesno("Umbenennen bestätigen", f"{kopf}?\n\n{vorschau}"):
            return

        ok = 0
        for item, src, alt, dst, neu, verschieben in plan:
            try:
                quelle = os.path.join(src, alt)
                if verschieben:
                    os.makedirs(dst, exist_ok=True)
                    shutil.move(quelle, os.path.join(dst, neu))
                else:
                    os.rename(quelle, os.path.join(dst, neu))
                self.rows[item]["_orig"] = neu
                self.rows[item]["_dir"] = dst
                hinweis = "✓ einsortiert" if verschieben else "✓ umbenannt"
                self.tree.item(item, values=(neu, self.rows[item].get("dokumententyp", ""),
                                             self.rows[item].get("datum", ""),
                                             self.rows[item].get("projekt", "") or "—",
                                             hinweis))
                self.protokoll(zeile(alt, dst, neu, verschieben))
                ok += 1
            except OSError as e:
                self.protokoll(f"FEHLER bei {alt}: {e}")
        self.status.set(f"{ok}/{len(plan)} Datei(en) verarbeitet.")
        messagebox.showinfo("Fertig", f"{ok} von {len(plan)} Datei(en) umbenannt"
                            + (f", {anzahl_verschoben} in Projektordner einsortiert." if anzahl_verschoben else "."))


def main():
    if tk is None:
        sys.stderr.write(
            "Tkinter ist nicht installiert.\n"
            "Linux:   sudo apt install python3-tk   (bzw. python3-tkinter)\n"
            "Windows/macOS: die offiziellen Python-Installer von python.org "
            "enthalten Tkinter bereits.\n")
        sys.exit(1)
    global _HAS_DND
    root = None
    if _HAS_DND:
        try:
            root = TkinterDnD.Tk()      # DnD-faehiges Fenster
        except Exception:
            _HAS_DND = False            # Fenster trotzdem oeffnen, ohne DnD
            root = None
    if root is None:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
