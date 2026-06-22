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

1. **Dokumente laden** – mehrere Wege:
   - **Dateien wählen…** – Mehrfachauswahl per Dialog (auch `.eml`/`.msg`).
     Funktioniert immer, auch wenn Drag & Drop nicht geht.
   - **Reinziehen:** Dateien, ganze Ordner oder **E-Mails** ins Fenster ziehen.
   - oder **Ordner** wählen → **Einlesen**.

   > **Outlook-Hinweis:** Eine E-Mail **direkt aus Outlook** ins Fenster zu
   > ziehen liefert oft keine Datei. Speichere sie zuerst als `.msg`
   > (in Outlook „Speichern unter" oder auf den Desktop ziehen) und nutze dann
   > „Dateien wählen…" bzw. Drag & Drop. Beim Start zeigt das Protokoll unten,
   > ob Drag & Drop aktiv ist.

   **E-Mails (`.eml`/`.msg`)** werden automatisch zerlegt: der **Mailtext wird
   als PDF** erzeugt und **alle Anhänge** als eigene Dateien herausgelöst –
   danach werden PDF und Anhänge wie normale Dokumente nach VA benannt.
   (`.msg` von Outlook nutzt `extract-msg`; in der `.exe` enthalten.)
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

## Automatisch in Projektordner einsortieren

Ist der Haken **„Umbenannte Dateien in Projektordner einsortieren"** gesetzt
(Standard: an), legt das Programm jede Datei nach dem Umbenennen direkt in den
passenden **Projektordner** unter der **Projektbasis**.

- Die **Projektbasis** ist der Ordner, der die Projekt-Unterordner enthält
  (Standard: `C:\Users\ziegler\Desktop\Dokumentenumbenennung`). Liest du den
  Übergabeordner `00_Posteingang` ein, wird die Projektbasis automatisch auf
  dessen übergeordneten Ordner gesetzt.
- Das Projekt wird aus **Dateiname + Inhalt/Mailtext** erkannt – am stärksten
  über die **Projektnummer** (z. B. `0875` / `0875.20`), sonst über Orts-/
  Stichworte aus dem Ordnernamen bzw. aus `projekte_mapping.json`. Alle aus
  **einer** E-Mail erzeugten Dateien (Mailtext-PDF + Anhänge) landen im selben
  Projektordner.
- Die Spalte **„Projektordner"** zeigt das erkannte Ziel. Stimmt es nicht oder
  steht dort `—` (unklar/mehrdeutig), trägst du im Feld **„Projektordner
  (Ablage)"** den richtigen Ordnernamen ein und klickst **Vorschau
  aktualisieren**.
- **Sicher:** Einsortiert wird **nur** in einen Ordner, der unter der
  Projektbasis bereits existiert. Wird kein eindeutiges Projekt erkannt, bleibt
  die Datei einfach im Ausgangsordner liegen (sie wird trotzdem umbenannt).
  Haken aus = nur umbenennen, nicht verschieben.

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
| `email_extract.py` | E-Mail-Zerlegung (Mailtext-PDF + Anhänge) |
| `projekt_zuordnung.py` | Projekt-Erkennung für die Einsortierung in Projektordner |
| `api_client.py` | Optionaler Claude-API-Modus |

> Die verbindlichen Regeln stehen zusätzlich in
> [`../.claude/skills/dokumentenumbenennung/SKILL.md`](../.claude/skills/dokumentenumbenennung/SKILL.md)
> (Quelle des API-Systemprompts) und in
> [`../docs/VA_Dokumentenbenennung_1.1.pdf`](../docs/VA_Dokumentenbenennung_1.1.pdf).
