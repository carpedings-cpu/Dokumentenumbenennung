# Dokumenten-Umbenenner (Desktop-Programm)

Eigenständiges Programm, das Projektdokumente nach der KPC-Verfahrensanweisung
*„Dokumentenbenennung & Versionierung"* (VA 1.1) umbenennt – **außerhalb von
Claude**, mit einer kleinen Fenster-Oberfläche.

## Zwei Modi (Umschalter)

| Modus | Was es tut | Voraussetzung |
|-------|-----------|---------------|
| **Offline (regelbasiert)** | Liest das **Datum** automatisch aus, schlägt einen **Typ** vor; du bestätigst/ergänzt die Felder. | Nur Python (Tkinter ist dabei). Kein Internet. |
| **Claude-API** | Erkennt Datum, Quelle, Typ, Bezeichnung **automatisch** – auch bei Scans. | API-Schlüssel + `pip install anthropic`. |

In beiden Fällen gilt: **Vorschau ansehen → bestätigen → umbenennen.** Es wird
nichts überschrieben (Konflikte bekommen `-02`, `-03`, …), die Dateiendung
bleibt erhalten.

## Starten

Voraussetzung: **Python 3.9+** (unter Windows „Add Python to PATH" anhaken;
Tkinter ist bei den offiziellen Installern dabei).

```bash
cd tool
python dokumenten_umbenenner.py
```

Unter Windows genügt meist ein Doppelklick auf `dokumenten_umbenenner.py`.

## Bedienung

1. **Dokumente laden** – zwei Wege:
   - **Reinziehen:** Dateien oder ganze Ordner direkt ins Fenster ziehen
     (Drag & Drop).
   - oder **Ordner** wählen → **Einlesen**.
2. Modus wählen:
   - **Offline:** erkennt das Datum automatisch und schlägt einen Typ vor.
   - **Claude-API** oder **Gemini-API:** erkennt **alle Felder automatisch**
     (auch bei Scans). Dazu den passenden Modus anklicken und oben den
     **API-Schlüssel** dieses Anbieters eintragen.
     - Claude-Schlüssel: console.anthropic.com → API Keys (`sk-ant-…`)
     - Gemini-Schlüssel: aistudio.google.com/apikey (`AIza…`, Modell
       `gemini-2.5-flash`)
3. In der Tabelle eine Datei anklicken, unten die **Felder** prüfen/ergänzen,
   **Vorschau aktualisieren**. Für Pläne den Haken *„Planunterlage"* setzen.
4. **Alle umbenennen** → Bestätigungsdialog → fertig.

> Drag & Drop nutzt das Paket `tkinterdnd2`; in der `.exe` ist es bereits
> enthalten. Beim Start aus dem Quellcode ggf. `pip install tkinterdnd2`
> (ohne das Paket funktioniert weiterhin der Weg über „Ordner wählen").

## Optional: API-Modus einrichten

```bash
pip install anthropic
```

Schlüssel ins Feld eintragen oder vorab als Umgebungsvariable setzen:

```bash
export ANTHROPIC_API_KEY="sk-ant-…"     # Windows: setx ANTHROPIC_API_KEY "sk-ant-…"
```

Der API-Modus nutzt das Modell `claude-opus-4-8` und schickt das PDF direkt zur
Analyse (kostenpflichtig pro Dokument).

## Als Windows-`.exe` (ohne Python)

Damit das Programm ganz ohne Python-Installation läuft, kann es zu einer
einzelnen `.exe` gebündelt werden. Zwei Wege:

### A) Automatisch über GitHub (kein Windows-Rechner nötig)
1. Im Repo auf den Reiter **Actions** gehen.
2. Links **„Windows-EXE bauen"** wählen → rechts **„Run workflow"**.
3. Nach ein paar Minuten unten beim Lauf unter **Artifacts** die
   `Dokumenten-Umbenenner-Windows` herunterladen → entpacken → `.exe` starten.

Alternativ: einen Versions-Tag setzen (z. B. `v1.0`) – dann wird automatisch ein
**Release** mit der `.exe` zum Download erstellt.

### B) Lokal auf einem Windows-Rechner
Mit installiertem Python: Doppelklick auf [`build_windows.bat`](build_windows.bat)
(oder im Repo-Stammverzeichnis `pyinstaller --noconfirm dokumenten-umbenenner.spec`).
Die fertige Datei liegt danach unter `dist/Dokumenten-Umbenenner.exe`.

Die `.exe` enthält **beide Modi** – offline funktioniert sofort, der API-Modus
braucht zusätzlich einen Schlüssel im Programm.

## Dateien

| Datei | Zweck |
|-------|-------|
| `dokumenten_umbenenner.py` | Programm mit Oberfläche (Start hier) |
| `va_rules.py` | VA-Regeln: Typliste, Datums-/Namenslogik, Bereinigung |
| `pdf_text.py` | PDF-Textextraktion ohne Zusatzpakete |
| `api_client.py` | Optionaler Claude-API-Modus |

> Die verbindlichen Regeln stehen zusätzlich in
> [`../.claude/skills/dokumentenumbenennung/SKILL.md`](../.claude/skills/dokumentenumbenennung/SKILL.md)
> (Quelle des API-Systemprompts) und in
> [`../docs/VA_Dokumentenbenennung_1.1.pdf`](../docs/VA_Dokumentenbenennung_1.1.pdf).
