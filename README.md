# Dokumentenumbenennung

Tool zur **einheitlichen Benennung von Projektdokumenten** nach der
KPC-Verfahrensanweisung *„Dokumentenbenennung & Versionierung"* (VA, Version 1.1).

Umgesetzt als **Claude Code Skill** – keine Installation, keine Abhängigkeiten.
Claude liest das Dokument, bildet den Dateinamen nach Schema, zeigt eine
Vorschau und benennt nach Bestätigung um.

## Verwendung

In Claude Code im Projektordner z. B.:

- „Benenne die Dateien im Ordner `./Eingang` nach der VA um."
- „Wie muss dieses Protokoll vom 15.04. heißen?"
- „Bereinige die Dokumentenbenennung in diesem Verzeichnis."

Claude erkennt die passende Skill automatisch (`dokumentenumbenennung`) und
arbeitet nach diesem Ablauf:

1. **Lesen** – Inhalt/Metadaten des Dokuments ermitteln.
2. **Erkennen** – Datum, Quelle, Phase, Dokumententyp, Bezeichnung, Version.
3. **Vorschau** – Tabelle `alt → neu` (Dry-Run), bevor etwas passiert.
4. **Umbenennen** – nach Bestätigung (`git mv`/`mv`, Konflikte werden vermieden).

## Schema (Kurzfassung)

```
JJMMTT_[Quelle]_[Phase]_Dokumententyp_Bezeichnung[_Version]
```

Planunterlagen folgen einem eigenen Schema:

```
[Projektnr.].[Auftragsart]_[Leistungsphase]_[Planinhalt]_[Geschoss]_[Plannr.]_[Index]_[Status]
```

Die vollständigen Regeln, Datums-Ausnahmen, die verbindliche Typ-Referenzliste
und das Planbenennungsschema stehen in:

- [`.claude/skills/dokumentenumbenennung/SKILL.md`](.claude/skills/dokumentenumbenennung/SKILL.md)

## Grundregeln

- Keine Leerzeichen/Sonderzeichen; Umlaute erlaubt.
- `_` trennt Schema-Felder, `-` verbindet Qualifizierer innerhalb eines Feldes.
- Dokumententyp **nur** aus der Referenzliste.
- Datum entfällt u. a. bei Plan-, Vertrags- und extern erstellten Dokumenten.
