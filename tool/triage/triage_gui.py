#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auswahl-Fenster fuer die KPC-Triage.

Zeigt die neuen Posteingangs-Mails als Liste mit Haekchen. Du klickst die Mails
an, die umbenannt werden sollen; nur diese werden als .msg an den
Dokumentenbenennungs-Skill uebergeben. Outlook wird ausschliesslich gelesen.

Start ueber 3_Mails_auswaehlen.bat (bzw. 3b_Heute_auswaehlen.bat).
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except ModuleNotFoundError:
    tk = None

import triage as T  # noqa: E402


class AuswahlFenster:
    def __init__(self, root, ctx, zeilen, neue_max, mit_kategorie):
        self.root = root
        self.ctx = ctx
        self.zeilen = zeilen
        self.neue_max = neue_max
        self.mit_kategorie = mit_kategorie
        self.checked = {}     # iid -> bool
        self.row_of = {}      # iid -> zeile

        root.title("KPC Triage – Mails zum Umbenennen auswählen")
        root.geometry("1120x650")

        kopf = ttk.Frame(root, padding=8)
        kopf.pack(fill="x")
        ttk.Label(kopf, text=(f"{len(zeilen)} neue Mail(s). Häkchen setzen bei den Mails, die "
                              "umbenannt werden sollen - dann unten auf 'Ausgewählte ablegen'.")
                  ).pack(side="left")

        rahmen = ttk.Frame(root, padding=(8, 0))
        rahmen.pack(fill="both", expand=True)
        cols = ("sel", "dringlichkeit", "projekt", "kategorie", "absender", "betreff", "eingang")
        self.tree = ttk.Treeview(rahmen, columns=cols, show="headings", height=20)
        for c, t, w in (("sel", "✓", 40), ("dringlichkeit", "Dringl.", 70),
                        ("projekt", "Projekt", 170), ("kategorie", "Kategorie", 160),
                        ("absender", "Absender", 200), ("betreff", "Betreff", 320),
                        ("eingang", "Eingang", 120)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(rahmen, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self.klick)

        for z in zeilen:
            vorab = bool(z.get("_relevant"))   # Wichtige sind vorausgewählt
            iid = self.tree.insert("", "end", values=(
                "☑" if vorab else "☐", z["gruppe"],
                (str(z["kuerzel"]) + " " + str(z["projekt"])).strip(),
                z["kategorie"] + (" (KI)" if z.get("ki") else ""),
                z["absender"], z["betreff"], z["eingang"]))
            self.checked[iid] = vorab
            self.row_of[iid] = z

        leiste = ttk.Frame(root, padding=8)
        leiste.pack(fill="x")
        ttk.Button(leiste, text="Alle", command=lambda: self.setze_alle(True)).pack(side="left")
        ttk.Button(leiste, text="Keine", command=lambda: self.setze_alle(False)).pack(side="left", padx=4)
        ttk.Button(leiste, text="Ausgewählte ablegen", command=self.ablegen).pack(side="right")
        self.status = tk.StringVar(value="Tipp: Zeile anklicken = Häkchen umschalten.")
        ttk.Label(leiste, textvariable=self.status).pack(side="right", padx=12)

    # -- Interaktion --
    def klick(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        if iid:
            self._setze(iid, not self.checked.get(iid, False))

    def _setze(self, iid, wert):
        self.checked[iid] = wert
        vals = list(self.tree.item(iid, "values"))
        vals[0] = "☑" if wert else "☐"
        self.tree.item(iid, values=vals)

    def setze_alle(self, wert):
        for iid in self.tree.get_children():
            self._setze(iid, wert)

    def ablegen(self):
        auswahl = [self.row_of[iid] for iid in self.tree.get_children()
                   if self.checked.get(iid)]
        if not auswahl:
            messagebox.showinfo("Nichts ausgewählt", "Bitte mindestens eine Mail anhaken.")
            return
        if not messagebox.askyesno(
                "Bestätigen",
                f"{len(auswahl)} Mail(s) ins Körbchen legen?\n\n"
                "Die übrigen Mails gelten danach als erledigt und werden beim nächsten "
                "Lauf nicht erneut gezeigt ('Heute auswählen' holt sie zurück)."):
            return
        try:
            n = T.exportiere(auswahl, self.ctx["eingang"], self.mit_kategorie, self.ctx["cfg"])
            T.marker_weiterstellen(self.ctx, self.neue_max, self.zeilen)
            T.schreibe_bericht(auswahl, self.ctx["bericht_dir"], True)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Fehler", str(e))
            return
        messagebox.showinfo(
            "Fertig",
            f"{n} Mail(s) als .msg abgelegt in:\n{self.ctx['eingang']}\n\n"
            "Jetzt das Umbenenn-Programm auf diesen Ordner anwenden "
            "(Ordner wählen → Einlesen).")
        self.root.destroy()


def main():
    if tk is None:
        sys.stderr.write("Tkinter fehlt. Unter Windows ist es im Python-Installer enthalten.\n")
        sys.exit(1)
    ap = argparse.ArgumentParser(description="KPC Triage – Auswahl-Fenster.")
    ap.add_argument("--heute", action="store_true", help="Heutige Mails erneut zur Auswahl zeigen.")
    ap.add_argument("--seit", default=None, help="Startdatum YYYY-MM-DD.")
    ap.add_argument("--stufe2", action="store_true", help="Stufe-2-API erzwingen.")
    ap.add_argument("--mit-kategorie", action="store_true",
                    help="Reversible Outlook-Kategorie setzen.")
    args = ap.parse_args()

    ctx = T.vorbereiten(heute=args.heute, seit=args.seit, stufe2=args.stufe2)

    root = tk.Tk()
    root.withdraw()
    try:
        zeilen, neue_max = T.sammle_mails(ctx)
    except Exception as e:  # noqa: BLE001
        messagebox.showerror("Outlook", f"Konnte Outlook nicht lesen:\n{e}\n\n"
                             "Bitte klassisches Desktop-Outlook öffnen und erneut versuchen.")
        return
    if not zeilen:
        messagebox.showinfo("Keine neuen Mails",
                            "Es gibt keine neuen Mails seit dem letzten Lauf.\n\n"
                            "Tipp: '3b Heute auswaehlen' zeigt die heutigen erneut.")
        return
    root.deiconify()
    AuswahlFenster(root, ctx, zeilen, neue_max, args.mit_kategorie)
    root.mainloop()


if __name__ == "__main__":
    main()
