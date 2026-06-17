---
name: dokumentenumbenennung
description: >-
  Benennt Projektdokumente nach der KPC-Verfahrensanweisung "Dokumentenbenennung
  & Versionierung" (VA, Version 1.1) um. Nutze diese Skill, wenn der Nutzer
  Dateien/Scans/PDFs umbenennen, ein Dokument korrekt benennen, einen Dateinamen
  nach Schema bilden oder einen Ordner nach der VA bereinigen möchte. Erkennt
  Datum, Quelle, Phase, Dokumententyp, Bezeichnung und Version, schlägt den neuen
  Namen als Vorschau (Dry-Run) vor und benennt nach Bestätigung um.
---

# Dokumentenbenennung (KPC VA 1.1)

Ziel: Jeder Dateiname lässt **ohne Öffnen der Datei** erkennen, *wann* das
Dokument entstand, *von wem* es stammt (bei externen Dokumenten), *welcher
Dokumententyp* es ist und *worum* es inhaltlich geht.

## Verbindliches Schema

```
JJMMTT_[Quelle]_[Phase]_Dokumententyp_Bezeichnung[_Version]
```

| Feld          | Erläuterung |
|---------------|-------------|
| `JJMMTT`      | Datum der Erstellung, des Eingangs oder des Versands (z. B. `250415`) |
| `[Quelle]`    | Nur bei **extern** erstellten Dokumenten: Kunde, Planer, Lieferant, Behörde |
| `[Phase]`     | Nur bei **Abnahme / Einweisung / IBN** |
| `Dokumententyp` | Festgelegter Typ **aus der Referenzliste** (siehe unten) – nur diese Begriffe sind zulässig |
| `Bezeichnung` | Kurze, präzise inhaltliche Beschreibung |
| `[_Version]`  | Nur bei überarbeiteten oder freigegebenen Dokumenten |

Eckige Klammern `[ ]` kennzeichnen **variable/optionale** Felder – sie stehen
nie im echten Dateinamen.

## Schreibregeln

- **Keine Leerzeichen, keine Sonderzeichen.** Umlaute (ä ö ü) und ß sind erlaubt.
- **Unterstrich `_`** trennt die Schema-Felder.
- **Bindestrich `-`** verbindet Qualifizierer **innerhalb** eines Feldes
  (z. B. `Baubesprechung-KW16`, `Nachtrag-03`, `Lieferant-A`, `Sanitär-EG`, `Typ-X`).
- Original-Dateiendung (`.pdf`, `.docx`, …) immer beibehalten.

## Datumsregeln – Ausnahmen (KEIN Datum)

Bei diesen Dokumenten entfällt das Datumsfeld `JJMMTT`:

- **Planunterlagen** → eigenes Schema (siehe „Planbenennungsschema").
- **Vertragsunterlagen** → Typ und Bezeichnung sind eindeutig.
- **Betriebsanleitungen & Wartungsvorgaben** → Gerät/Anlage als Bezeichnung.
- **Extern erstellte Dokumente** (Kapitel 5.8) → Quelle und Typ sind ausreichend.

## Versionierung

Die Versionsangabe steht am **Ende** des Dateinamens.

| Version       | Bedeutung |
|---------------|-----------|
| `0.1, 0.2, …` | Entwürfe / interne Arbeitsstände |
| `1.0`         | Erster freigegebener Stand – abgeschlossen und gesperrt |
| `1.1, 1.2, …` | Überarbeitete Versionen eines freigegebenen Dokuments |
| `2.0`         | Zweiter freigegebener Stand |
| `A, B, C …`   | Plan-Index – Überarbeitungsstand von Planunterlagen; ersetzt die vorherige Version |

> Ein freigegebener Stand (1.0 bzw. F bei Plänen oder höher) gilt als
> abgeschlossen. Änderungen nur nach erneuter fachlicher/kaufmännischer
> Bewertung; ein Nachtrag ist mit der Geschäftsführung abzustimmen.

## Beispiele

| Dateiname | Erläuterung |
|-----------|-------------|
| `250415_Vermerk_Baubegehung` | Intern, Vermerk (Datum = Erstellung) |
| `250415_Protokoll_Baubesprechung-KW16` | Intern, Protokoll (Datum = Termin) |
| `250415_Kunde_Schriftverkehr_Nachtrag-03` | Extern (Kunde), Schriftverkehr |
| `250415_IBN-Protokoll_Lieferant-A_Küchentechnik` | IBN-Protokoll, Lieferant A, Küchentechnik |
| `250415_Prüfprotokoll_Abnahme_Sanitär-EG` | Prüfprotokoll Abnahme, Gewerk Sanitär, EG |
| `250415_Abnahmeprotokoll_1.0` | Abnahmeprotokoll, freigegeben |
| `250415_Versandnachweis_Revisionsunterlagen` | Intern, Versandnachweis Übergabe |
| `250415_E-Mail_Kunde_Nachtragsbeauftragung` | E-Mail mit vertraglicher Relevanz, Kunde |
| `Betriebsanleitung_Kälteanlage-Typ-X` | Betriebsanleitung, kein Datum |

---

# Dokumententypen – Referenzliste (verbindlich)

Nur diese Typbezeichnungen dürfen im Dateinamen verwendet werden.

### 5.1 Allgemeine Typen (projektphasenübergreifend)
`Vermerk`, `Gesprächsnotiz`, `Protokoll`, `Schriftverkehr`, `E-Mail`,
`Bericht`, `Anweisung`, `Entscheidung`, `Versandnachweis`, `Datenblatt`,
`Steuerungsdokument`

### 5.2 Vertragliche Grundlagen / Vertragsmanagement
`Stückliste`, `Bauzeitenplan`, `Muster_[Bezeichnung]`, `Vorgabe_Montage`,
`Vorgabe_Revision`, `Vorgabe_Wartung`, `Vorgabe_Sicherheit`, `Vorgabe_Logistik`,
`Richtlinie`, `Baustellenordnung`, `Nachweisdokument`, `Prüfbericht`,
`Formblatt`, `Behinderungsanzeige`, `Vertrag_[Partnername]`,
`Vorgabe_[Bezeichnung]`
*(Vertragsunterlagen: kein Datum.)*

### 5.3 Planunterlagen
Eigenständiges Schema → siehe „Planbenennungsschema".

### 5.4 Inbetriebnahme (IBN)
`IBN_Protokoll_[Lieferant]_[Gewerk]`, `IBN_Anzeige`

### 5.5 Abnahme
`Abnahme_Anmeldung`, `Abnahme_Terminbestätigung`,
`Abnahme_Prüfprotokoll_[Gewerk]`, `Abnahmeprotokoll`,
`Abnahmeprotokoll_[Bereich]`, `Abnahme_Fertigstellungsanzeige`,
`Abnahme_Mängelliste`, `Abnahme_Stückliste`, `Abnahme_Freigabe`,
`Abnahme_Mängelfreimeldung`, `Abnahme_Leistungsfeststellung`

### 5.6 Betrieb, Service, Wartung
`Revisionsunterlagen`, `Betriebsanleitung_[Gerät]` (kein Datum),
`Wartungsvorgabe_[Anlage]` (kein Datum), `Wartungsvertrag`,
`Wartungsvertrag_[Partnername]`, `Servicevertrag_[Partnername]`,
`Wartungsanforderung`, `Abschlussanzeige`

### 5.8 Extern erstellte Dokumente (Quelle im Namen, KEIN Datum)
`Auftragsbestätigung_[Lieferant]`, `Lieferavis`, `Lieferterminänderung`,
`Freigabe`, `Stellungnahme`, `Bauzeitenplan`, `Werksplan_[Lieferant]`,
`Lieferantenplan_[Lieferant]`

---

# Planbenennungsschema (Kapitel 6)

Planunterlagen folgen einem eigenständigen Schema. **Das Datum entfällt** –
Projektnummer, Index und Status sind eindeutig.

```
[Projektnr.].[Auftragsart]_[Leistungsphase]_[Planinhalt]_[Geschoss]_[Plannr.]_[Index]_[Status]
```

**Beispiel:** `1030.20_8_E_EG_01_A_V`
*(Projekt 1030, Ausführung, Objektüberwachung, Einrichtung, Erdgeschoss,
Plan 01, Index A, Vorabzug)*

| Feld | Kodierung / Werte |
|------|-------------------|
| Projektnr. | Vierstellige Projektnummer aus Odoo |
| Auftragsart | `10` Planung · `20` Ausführung |
| Leistungsphase | `1` Grundlagenermittlung · `2` Vorplanung · `3` Entwurfsplanung · `4` Genehmigungsplanung · `5` Ausführungsplanung · `6` Vorbereitung der Vergabe · `7` Mitwirkung bei der Vergabe · `8` Objektüberwachung/Ausführung · `9` Objektbetreuung |
| Planinhalt | `A` Ansichten · `B` Bodenplan · `D` Details · `E` Einrichtung · `F` Flächenkonzept · `G` Gründung/Fundament · `I` Installation · `K` Kälte · `L` Lüftungsdecken/Hauben · `Ü` Übersichtsplan · `W` Wände (Rammschutz/Wandverstärkung) |
| Geschoss | `U2` 2. UG · `U1` 1. UG · `EG` Erdgeschoss · `01` 1. OG · `02` 2. OG · `03` 3. OG · `DG` Dachgeschoss · `99` geschossübergreifend |
| Plannr. (fortlfd.) | Zweistellig, je `Planinhalt × Geschoss`-Kombination **unabhängig** ab `01` gezählt. Bsp.: `A_U2_01`, `A_U2_02` → `B_U2_01` → `A_EG_01` → `E_EG_01` |
| Index | `A, B, C …` – Überarbeitungsstand des Plans |
| Status | `V` Vorabzug (nur Index A) · `PL` Prüflauf · `F` Freigegeben |

---

# Arbeitsweise (Workflow)

Beim Umbenennen **immer** so vorgehen:

1. **Eingabe klären.** Welche Datei(en) bzw. welcher Ordner? Wenn unklar,
   kurz nachfragen.
2. **Inhalt ermitteln.** Datei (oder den extrahierten Text) lesen. Datum aus
   Inhalt/Metadaten ableiten; bei externen Dokumenten die Quelle bestimmen.
3. **Felder festlegen.**
   - Datum → Format `JJMMTT` (sofern nicht ausgenommen).
   - Quelle/Phase nur setzen, wenn laut Schema erforderlich.
   - **Dokumententyp muss aus der Referenzliste stammen.** Passt nichts exakt,
     den nächstliegenden Typ vorschlagen und kurz rückfragen statt zu raten.
   - Bezeichnung kurz und präzise.
   - Version nur bei Entwürfen/freigegebenen Ständen.
   - Planunterlagen → Planbenennungsschema verwenden.
4. **Namen bauen & bereinigen.** Leerzeichen/Sonderzeichen entfernen, `_` als
   Feldtrenner, `-` innerhalb eines Feldes, Umlaute behalten, Endung anhängen.
5. **Vorschau (Dry-Run) zeigen.** Tabelle `alt → neu` ausgeben, **bevor** etwas
   umbenannt wird. Bei mehreren Dateien alle Vorschläge auf einmal zeigen.
6. **Nach Bestätigung umbenennen.** In Git-Repos `git mv`, sonst `mv`. Bei
   Namenskonflikten Suffix `-02`, `-03`, … anhängen. Nichts überschreiben.

**Unsicherheit immer transparent machen** und lieber kurz rückfragen, als einen
falschen Typ oder ein falsches Datum zu erfinden. Bei Fragen zur Benennung
einzelner Dokumente ist laut VA die Projektleitung zuständig.
