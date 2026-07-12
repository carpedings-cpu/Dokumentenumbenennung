# Dokumenten-Umbenenner – Kurzanleitung

Benennt Dokumente, Scans und E-Mails automatisch nach der KPC-Verfahrensanweisung
*„Dokumentenbenennung & Versionierung"* um und sortiert sie in die Projektordner.
**Kein Zusatzprogramm nötig** – einfach die `.exe` starten.

---

## 1. Starten
1. **`Dokumenten-Umbenenner.exe`** doppelklicken.
2. Windows meldet evtl. *„Der Computer wurde geschützt"* → **„Weitere Informationen"**
   → **„Trotzdem ausführen"**. (Nur beim ersten Mal.)

## 2. Dokumente laden (ein Weg genügt)
- **„Dateien wählen…"** anklicken und Dateien auswählen, **oder**
- Dateien/E-Mails einfach **ins Fenster ziehen**, **oder**
- oben einen **Ordner** wählen → **„Einlesen"**.

> E-Mails (`.eml`/`.msg`) werden automatisch zerlegt: **Mailtext als PDF** +
> **echte Anhänge** (Signatur-Bildchen werden weggelassen).

## 3. Prüfen und umbenennen
1. In der Tabelle siehst du je Datei **„Neu (Vorschau)"** (neuer Name) und
   **„Projektordner"** (wohin sie kommt).
2. Stimmt etwas nicht: Zeile anklicken → unten die **Felder** korrigieren →
   **„Vorschau aktualisieren"**.
3. Unten **„Alle umbenennen"** → bestätigen. **Fertig.**

## 4. In Projektordner einsortieren (optional, standardmäßig an)
- **Projektbasis** = der Ordner, in dem deine Projektordner liegen. Liest du den
  Übergabeordner `00_Posteingang` ein, stellt sich das automatisch richtig ein –
  sonst per **„Durchsuchen…"** deinen Projektordner wählen.
- Haken **„fehlende Projektordner anlegen"**: legt einen fehlenden Projektordner
  automatisch an (Name aus der Projektliste).

## Offline oder mit KI-Erkennung
- **Offline** (Standard): erkennt **Datum** und **Typ** automatisch, den Rest
  bestätigst du. Kein Internet, kein Schlüssel nötig.
- **Claude-API / Gemini-API**: erkennt **alle Felder automatisch** (auch bei
  Scans). Dazu oben den Modus anklicken und einen **eigenen API-Schlüssel**
  eintragen.

---

**Gut zu wissen:** Es wird **nichts überschrieben** (bei gleichem Namen `-02`,
`-03`, …), die **Dateiendung bleibt** erhalten. Nichts wird ins Internet geladen
(außer im API-Modus das jeweilige Dokument zur Erkennung).

*Fragen? → [hier deinen Namen/Kontakt eintragen]*
