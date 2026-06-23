# KPC Morgenbriefing (Outlook, Windows)

Liest morgens deinen Outlook-**Posteingang + Gesendete** (nur lesend), fasst die
Mails mit **Google Gemini** zu einem **Briefing** zusammen, erkennt **Termine/
Fristen** und schreibt eine **HTML-Übersicht** (öffnet sich automatisch) sowie ein
**laufendes Protokoll** (`briefing_log.md`, das „Gedächtnis"). Erkannte Termine
werden dir in einem Fenster zur Auswahl angeboten und **nur nach Bestätigung** in
den Outlook-Kalender eingetragen.

## Sicherheit / Datenschutz
- **E-Mails werden nur gelesen.** Einzige Schreibaktion: die von dir bestätigten
  Kalender-Termine anlegen.
- **Inkrementell:** nur Mails seit dem letzten Lauf (Marker in `briefing_state.json`).
- Für die Zusammenfassung gehen **Betreff + Textauszug** an die Gemini-API.

## Voraussetzungen
- Klassisches **Desktop-Outlook** (nicht „neues Outlook"/Store-App – die hat kein COM).
- **Python 3** und das Paket **pywin32** (installiert `0_Einrichten.bat`).
- Ein **Gemini-Schlüssel**: https://aistudio.google.com/apikey

## Einrichten (einmalig)
1. **`0_Einrichten.bat`** doppelklicken (installiert pywin32, legt `.env` an).
2. **`.env`** öffnen und den Schlüssel eintragen:
   `GEMINI_API_KEY=AIza...`

## Benutzen
- **`Briefing.bat`** doppelklicken → Briefing öffnet sich → im Termin-Fenster
  anhaken, was in den Kalender soll → „Ausgewählte in Kalender eintragen".

### Automatisch jeden Morgen (optional)
Windows **Aufgabenplanung** → „Einfache Aufgabe erstellen" → täglich z. B. 7:30 →
Programm: `Briefing.bat` (Pfad zu dieser Datei).

## Für Technische: Befehle
```
python briefing.py                 # seit letztem Lauf
python briefing.py --stunden 48    # Rueckblick 48 Stunden
python briefing.py --seit 2026-06-20
python briefing.py --kein-kalender # nur Briefing, kein Termin-Fenster
```

## Dateien
| Datei | Zweck |
|-------|-------|
| `briefing.py` | Hauptprogramm (read-only Outlook, Gemini, Kalender) |
| `.env` / `.env.example` | Gemini-Schlüssel |
| `briefing_state.json` | Marker des letzten Laufs (nicht eingecheckt) |
| `briefing_log.md` | laufendes Protokoll / Gedächtnis (nicht eingecheckt) |
| `briefing_config.json` | optionale Einstellungen (base_dir, max_mails …) |
