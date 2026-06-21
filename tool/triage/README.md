# KPC Posteingangs-Triage (Outlook, Windows)

Liest den Outlook-Posteingang **read-only**, klassifiziert eingehende Mails und
übergibt die **relevanten** Mails als `.msg` an den Dokumentenbenennungs-Skill.
Die Ablage/Benennung (Mailtext-PDF + Anhänge, Projekt-Unterordner) macht
weiterhin der Skill – die Triage baut das **nicht** doppelt.

## Sicherheit
- **Posteingang READ-ONLY:** kein Verschieben, Löschen oder Als-gelesen-markieren.
  Eigenschaften lesen und `.msg` exportieren ändern den Posteingang nicht.
- Einzige optionale Schreibaktion: eine **reversible Outlook-Kategorie**
  (`--mit-kategorie`, nur im scharfen Lauf).
- **Inkrementell:** nur Mails seit dem letzten *scharfen* Lauf (Marker in
  `triage_state.json`).
- **DSGVO:** lokal eindeutige Mails gehen **nie** an eine API. Stufe 2 (Claude)
  ist optional, standardmäßig **aus**, und sieht nur lokal unklare Mails.

## Einrichtung (einmalig)
```
pip install pywin32 anthropic
copy .env.example .env        # und ausfüllen (nur falls Stufe 2 gewünscht)
```
Outlook muss installiert und mit dem Konto eingerichtet sein (Desktop-Outlook,
nicht „neues Outlook"/Store-App – die hat kein COM).

## Ganz einfach: zum Doppelklicken (kein Tippen nötig)
Im Ordner `tool/triage/` liegen drei Dateien:

1. **`0_Einrichten.bat`** – einmal doppelklicken (installiert die nötigen Teile).
2. **`1_Probelauf.bat`** – Probelauf: schaut nur und öffnet die Übersicht,
   legt **nichts** ab und verändert **nichts**.
3. **`2_Scharf_schalten.bat`** – legt die wichtigen Mails wirklich ins
   Eingangs-Körbchen (fragt vorher zur Sicherheit nach).

> Beim ersten Doppelklick warnt Windows evtl. („Geschützt") → „Weitere
> Informationen" → „Trotzdem ausführen".

---

## Für Technische: Befehle
```
python triage.py                     # 1) TROCKENLAUF: nur klassifizieren + HTML
python triage.py --scharf            # 2) scharf: relevante Mails als .msg ablegen
python triage.py --scharf --mit-kategorie
python triage.py --seit 2026-06-01   # Marker einmalig ab Datum
python triage.py --stufe2            # Stufe-2-API erzwingen (sonst aus .env)
```
**Empfohlen:** zuerst Trockenlauf, HTML-Übersicht prüfen, dann scharf schalten.

## Konfiguration
- `triage_config.json` (optional) überschreibt die Standardwerte, u. a.:
  - `base_dir` – Projektbasis (Default `C:\Users\ziegler\Desktop\Dokumentenumbenennung`)
  - `eingang_unterordner` – Übergabeordner an den Skill (Default `00_Posteingang`)
  - `bericht_unterordner` – Ablage der HTML-Übersichten (Default `Triage-Berichte`)
- `projekte_mapping.json` – Projektname, Kürzel, Absender-Domains, Betreff-Stichworte.
  Neue Projekt-Unterordner unter `base_dir` werden beim Lauf automatisch ergänzt
  (mit leeren Domains/Stichworten) – Feinpflege machst du selbst.

## Klassifizierung
- **Stufe 1 (lokal, regelbasiert):** Projekt aus Absender-Domain/Betreff-Stichworten;
  Kategorie aus Betreff/Textauszug. Kategorien:
  `Auftragsbestaetigung-mit-Terminaenderung`, `Lieferavis`, `Maengel/Behinderung`,
  `Nachtrag`, `Rechnung`, `Info`.
- **Stufe 2 (optional, Claude):** nur für lokal nicht eindeutige Mails.

## Ausgabe
- **HTML-Übersicht** (KPC-Design, A4 quer) unter `Triage-Berichte/`, gruppiert nach
  Dringlichkeit (Hoch/Mittel/Niedrig/Unklar), mit Projekt, Kategorie, Absender,
  Betreff, Eingang.
- **`.msg`** der relevanten Mails (alles außer reiner `Info`) im Eingangsordner
  des Skills – von dort übernimmt der Dokumentenbenennungs-Skill.

## Dateien
| Datei | Zweck |
|-------|-------|
| `triage.py` | Hauptskript (read-only Triage) |
| `projekte_mapping.json` | Projektliste/Regeln (selbst pflegen) |
| `.env` / `.env.example` | Schalter + API-Schlüssel für Stufe 2 |
| `triage_state.json` | Marker (wird automatisch angelegt, nicht eingecheckt) |
