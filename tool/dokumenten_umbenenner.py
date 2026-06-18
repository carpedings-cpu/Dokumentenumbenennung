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
from pdf_text import extract_pdf_text  # noqa: E402

LESBARE_TEXT_ENDUNGEN = {".pdf", ".txt", ".md"}
ANALYSIERBAR = LESBARE_TEXT_ENDUNGEN | {".png", ".jpg", ".jpeg", ".gif", ".webp"}


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
        self.rows = {}  # tree-item-id -> dict mit Feldern

        self._baue_oben()
        self._baue_tabelle()
        self._baue_details()
        self._baue_unten()
        self._aktiviere_dnd()

    def _aktiviere_dnd(self):
        """Registriert das Fenster fuer Drag & Drop von Dateien/Ordnern."""
        if not _HAS_DND:
            return
        for ziel in (self.root, self.tree):
            try:
                ziel.drop_target_register(DND_FILES)
                ziel.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass

    # ----------------------------------------------------------------- UI
    def _baue_oben(self):
        f = ttk.Frame(self.root, padding=8)
        f.pack(fill="x")
        ttk.Label(f, text="Ordner:").pack(side="left")
        ttk.Entry(f, textvariable=self.ordner, width=70).pack(side="left", padx=4)
        ttk.Button(f, text="Durchsuchen…", command=self.waehle_ordner).pack(side="left")
        ttk.Button(f, text="Einlesen", command=self.einlesen).pack(side="left", padx=4)
        ttk.Button(f, text="Liste leeren", command=self._leeren).pack(side="left")

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

    def _baue_tabelle(self):
        hinweis = ("Dateien, ganze Ordner oder E-Mails (.eml/.msg) ins Fenster ziehen "
                   "- E-Mails werden in Mailtext-PDF + Anhänge zerlegt.") if _HAS_DND else \
                  ("Ordner waehlen und 'Einlesen'. "
                   "(Drag & Drop in dieser Version nicht verfuegbar.)")
        ttk.Label(self.root, text=hinweis).pack(anchor="w", padx=10, pady=(0, 2))

        f = ttk.Frame(self.root, padding=8)
        f.pack(fill="both", expand=True)
        cols = ("alt", "typ", "datum", "neu")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=12)
        for c, t, w in (("alt", "Alt (Ist)", 320), ("typ", "Typ", 150),
                        ("datum", "Datum", 80), ("neu", "Neu (Vorschau)", 420)):
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
        self._add_files([os.path.join(d, n) for n in self._dateien()], anhaengen=False)

    def on_drop(self, event):
        """Wird beim Hineinziehen von Dateien/Ordnern ausgelöst."""
        try:
            pfade = list(self.root.tk.splitlist(event.data))
        except Exception:
            pfade = [event.data]
        self._add_files(pfade, anhaengen=True)

    def _emails_extrahieren(self, dateien):
        """E-Mails (.eml/.msg) -> Mailtext-PDF + Anhaenge; sonst Datei unveraendert."""
        ergebnis = []
        for fp in dateien:
            if email_extract.ist_email(fp):
                try:
                    neu = email_extract.extrahiere(fp)
                    if neu:
                        self.protokoll(f"E-Mail {os.path.basename(fp)} -> {len(neu)} Datei(en) "
                                       "extrahiert (Mailtext + Anhänge).")
                        ergebnis += neu
                    else:
                        self.protokoll(f"  {os.path.basename(fp)}: nichts extrahierbar.")
                except Exception as e:  # noqa: BLE001
                    self.protokoll(f"  E-Mail-Fehler bei {os.path.basename(fp)}: {e}")
            else:
                ergebnis.append(fp)
        return ergebnis

    def _add_files(self, pfade, anhaengen=True):
        dateien = self._emails_extrahieren(self._expandiere(pfade))
        if not dateien:
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
        self.protokoll(f"Lese {len(dateien)} Datei(en) – Modus: {modus} …")
        for idx, fp in enumerate(dateien, 1):
            name = os.path.basename(fp)
            self.status.set(f"Analysiere {idx}/{len(dateien)}: {name}")
            self.root.update_idletasks()
            felder = self._analysiere(fp, name, modus, va_regeln)
            self._zeile_einfuegen(fp, felder)
        self.status.set(f"Fertig: {len(dateien)} Datei(en). Felder pruefen, dann 'Alle umbenennen'.")

    def _analysiere(self, pfad, name, modus, va_regeln):
        endung = os.path.splitext(name)[1].lower()
        felder = {k: "" for k in self.var}
        felder["ist_plan"] = False

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
        felder["datum"] = VA.datum_aus_text(text)
        felder["dokumententyp"] = VA.typ_aus_text(text)
        if felder["dokumententyp"] in VA.TYPEN_OHNE_DATUM:
            felder["datum"] = ""
        return felder

    def _stamm(self, felder):
        if felder.get("ist_plan"):
            return VA.baue_planname(felder)
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
                                        felder.get("datum", ""), neu_anzeige))
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
                                     felder.get("datum", ""), anzeige))

    def umbenennen(self):
        if not self.rows:
            return
        plan = []
        belegt = {}  # Ordner -> Set bereits geplanter Namen (Kollisionsschutz)
        for item, felder in self.rows.items():
            stamm = self._stamm(felder)
            if not stamm:
                continue
            d = felder["_dir"]
            vorhandene = belegt.setdefault(d, set())
            ziel = VA.eindeutiger_zielname(d, stamm, felder["_ext"],
                                           felder["_orig"], vorhandene=vorhandene)
            if ziel != felder["_orig"]:
                plan.append((item, d, felder["_orig"], ziel))
                vorhandene.add(ziel.lower())

        if not plan:
            messagebox.showinfo("Nichts zu tun",
                                "Keine Änderungen (Felder unvollständig oder Namen gleich).")
            return

        vorschau = "\n".join(f"{alt}  →  {neu}" for _, _, alt, neu in plan[:25])
        if len(plan) > 25:
            vorschau += f"\n… und {len(plan) - 25} weitere"
        if not messagebox.askyesno("Umbenennen bestätigen",
                                    f"{len(plan)} Datei(en) umbenennen?\n\n{vorschau}"):
            return

        ok = 0
        for item, d, alt, neu in plan:
            try:
                os.rename(os.path.join(d, alt), os.path.join(d, neu))
                self.tree.item(item, values=(neu, self.rows[item].get("dokumententyp", ""),
                                             self.rows[item].get("datum", ""), "✓ umbenannt"))
                self.rows[item]["_orig"] = neu
                self.protokoll(f"{alt}  →  {neu}")
                ok += 1
            except OSError as e:
                self.protokoll(f"FEHLER bei {alt}: {e}")
        self.status.set(f"{ok}/{len(plan)} Datei(en) umbenannt.")
        messagebox.showinfo("Fertig", f"{ok} von {len(plan)} Datei(en) umbenannt.")


def main():
    if tk is None:
        sys.stderr.write(
            "Tkinter ist nicht installiert.\n"
            "Linux:   sudo apt install python3-tk   (bzw. python3-tkinter)\n"
            "Windows/macOS: die offiziellen Python-Installer von python.org "
            "enthalten Tkinter bereits.\n")
        sys.exit(1)
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
