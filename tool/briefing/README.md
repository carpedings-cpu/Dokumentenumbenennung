# KPC Morgenbriefing (Outlook, Windows)

Liest morgens deinen Outlook-**Posteingang + Gesendete** (nur lesend), fasst die
Mails mit **Google Gemini** zu einem **Briefing** zusammen und schreibt eine
**HTML-Übersicht** (öffnet sich automatisch) plus ein **laufendes Protokoll**
(`briefing_log.md`).

**Aufgaben (To-dos):** Aus den Mails wird eine fortlaufende Aufgabenliste gepflegt
(`todos.json`). In einem Fenster kannst du Aufgaben **abhaken** – erledigte werden
gemerkt und **kommen nicht wieder**.

**Für wen?** Stehst du im **An**, gilt es als deine Aufgabe; bist du nur in
**Kopie/CC**, wird es separat als *„jemand anderes zuständig"* ausgewiesen.

**Termine:** werden erkannt, **mit Projekt** in den Outlook-Kalender eingetragen
(Betreff = „Projekt – Titel") und **nur nach Bestätigung**. **Terminänderungen**
(gleicher Termin, neues Datum) werden als *„GEÄNDERT: war … → jetzt …"* markiert.

## Sicherheit / Datenschutz
- **E-Mails werden nur gelesen.** Einzige Schreibaktion: die von dir bestätigten
  Kalender-Termine anlegen.
- **Inkrementell:** nur Mails seit dem letzten Lauf (Marker in `briefing_state.json`).
- Für die Zusammenfassung gehen **Betreff + Textauszug** an die Gemini-API.

## Voraussetzungen
- Klassisches **Desktop-Outlook** (nicht „neues Outlook"/Store-App – die hat kein COM).
- Ein **Gemini-Schlüssel**: https://aistudio.google.com/apikey

## Variante A (empfohlen): die `.exe` – kein Python nötig
1. **`KPC-Morgenbriefing.exe`** aus den GitHub-Releases herunterladen und in einen
   Ordner legen (z. B. auf den Desktop).
2. Einmal doppelklicken. Beim ersten Start wird daneben eine **`.env`** angelegt
   und du wirst gefragt, den Schlüssel einzutragen.
3. Die **`.env`** (liegt neben der `.exe`) öffnen → `GEMINI_API_KEY=AIza...`
   eintragen → speichern.
4. `KPC-Morgenbriefing.exe` erneut doppelklicken → Briefing öffnet sich → im
   Termin-Fenster anhaken, was in den Kalender soll → „Ausgewählte in Kalender
   eintragen".

> Wichtig: `.env`, das Protokoll und der Status liegen **neben der `.exe`** –
> die `.exe` also nicht in einen schreibgeschützten Ordner legen.

## Variante B: aus dem Quellcode (Python)
1. **`0_Einrichten.bat`** doppelklicken (installiert pywin32, legt `.env` an).
2. **`.env`** öffnen und den Schlüssel eintragen: `GEMINI_API_KEY=AIza...`
3. **`Briefing.bat`** doppelklicken.

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
