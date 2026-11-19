# Schulaktivitäten-Planer (Exkursionen & schulische Aktivitäten)

**Stand:** 2026-11-06  
**Status:** Phase 1 (Feedback Schulen)  
**Vorbild:** [`schularbeiten-planer.md`](./schularbeiten-planer.md)  
**Ziel:** Lehrer:innen beantragen Exkursionen / schulische Aktivitäten; Direktion/Admin genehmigt, fordert nach oder lehnt ab. Administration markiert Untis-Erledigung. Übersichten als Liste + Kalender, Export als ICS. **System of Record:** SharePoint-Listen auf der Intranet-Site.

---

## 1. Zielbild

```text
Lehrer / Direktion ──MSAL──► Schulaktivitäten-Planer
                                 │ Graph (Sites.ReadWrite.All)
                                 ▼
                        SharePoint Intranet-Site
                        ├─ Schulaktivitaeten   (Anträge)
                        └─ Aktivitaet-Regelwerk (optional, Seed)
                                 │ optional später
                                 ▼
                        Schulkalender / Outlook (ICS) / Schultermine-Sync
```

---

## 2. Entscheidungen (fest)

| Thema | Entscheidung |
|-------|--------------|
| UI-Stack | Vanilla/ESM, Look wie Schularbeiten-Planer (`akt-*` CSS) |
| Datenhaltung | SharePoint-Listen auf Intranet-Site |
| Stammdaten | Klassen + Lehrer aus `tenant-settings` / `app-data-v2` |
| Referenzen | `KlasseCode`, `LehrerCode` (+ E-Mail), keine Lookups |
| Statusfluss | `entwurf` → optional `kv_stellungnahme` → `beantragt` → `genehmigt` \| `abgelehnt` \| `nachforderung`; nach Genehmigung optional `abgesagt` |
| Genehmigungspfad | Regelwerk: `direktion` oder `kv_dann_direktion` (Freistadt) |
| Vorlauf | Tage + Modus `sperre` \| `warnung` \| `ausnahme` |
| Teilnahme | Klasse(n) / Unterrichtsgruppe / Einzelpersonen + Nichtteilnehmende; Typ u. a. `Klassenabwesenheit` |
| Admin-Abschluss | Kennzeichen `UntisErledigt` getrennt von Genehmigung |
| Rollen (MVP) | Demo-Umschalter Lehrer / KV / Admin; später Entra-Gruppen |
| Backend | Kein eigenes Azure; Graph wie Schularbeiten |
| ICS | Download der Filteransicht **und** abonnierbarer Feed (nur `genehmigt`) auf der Site-Dokumentbibliothek |
| Outlook pro Aktivität | Bei Status `genehmigt`: ICS-Download, Eintrag in den persönlichen Outlook-Kalender, optional Einladung an Klassen-M365-Gruppe (`M365GroupMail` / Klassen-Teams) bzw. Fallback Schüler-E-Mails + Begleitung (`Calendars.ReadWrite`, `Group.Read.All`) |
| Schultermine-Sync | optional später (Flag wie beim SA-Planer) |

**Power Automate:** Setup-Assistent `tools/schulaktivitaeten-setup.html` erzeugt ein Legacy-ZIP (Microsoft Approvals → Direktion, Status-Mail an Lehrkraft). Vorlage: `assets/power-automate/schulaktivitaeten/` (Generator `scripts/build-akt-pa-template.mjs`).

**Phase 2:** Unterrichtseinheiten (`StartUe`/`EndUe`) mit Standard-Zeitraster → abgeleitete Uhrzeiten; View **Berichte** (Jahresliste + Statistik nach Typ/Status/Monat/Klasse). ICS nutzt bei eintägigen Terminen die abgeleiteten Uhrzeiten.

**ICS-Feed (Abo):** Button **Feed** schreibt `MS365-Schultools/kalender/schulaktivitaeten.ics` per Graph in die Site-Bibliothek (Details in `DESCRIPTION` / `X-ALT-DESC`). Beim Veröffentlichen: anonyme Freigabe wenn erlaubt; sonst Org-`download.aspx` + Hinweis. **Outlook:** HTTPS-Link → Kalender hinzufügen → Aus dem Internet (nicht `webcal://` auf SharePoint-Pfaden). **Apple:** webcal. Ohne öffentliches Abo: ICS-Download oder „In meinen Kalender“. Button **ICS** = einmaliger Download der Filteransicht.

**Noch offen:** Schulspezifisches Zeitraster in Stammdaten, Wochenvorschau-Export, Untis-Import von Schülergruppen, Flow-Trigger auch bei Statuswechsel `nachforderung` → `beantragt`.

---

## 3. SharePoint-Listen

### 3.1 `Schulaktivitaeten`

| Interner Name | Display | Typ | Bemerkung |
|---------------|---------|-----|-----------|
| Title | Titel | Standard | z. B. Museumsbesuch |
| AktivitaetId | Aktivitäts-ID | text | `akt-` + kurz |
| Typ | Typ | choice | `Exkursion`, `Schulaktivitaet`, `Veranstaltung`, `Klassenabwesenheit`, `Sonstiges` |
| KlasseCode | Klasse-Code | text | Primärklasse (erste Auswahl) |
| KlasseCodes | Klassen-Codes | text | Mehrfach `;`-getrennt |
| TeilnahmeModus | Teilnahme-Modus | choice | `klasse`, `klassen`, `gruppe`, `schueler` |
| GruppenText | Unterrichtsgruppe | text | z. B. Religion über Klassen |
| TeilnehmendeText | Teilnehmende | multiline | Einzelpersonen |
| NichtTeilnehmende | Nicht teilnehmend | multiline | „Daheimbleiber“ |
| LehrerCode | Lehrer-Kürzel | text | Antragsteller |
| LehrerEmail | Lehrer-E-Mail | text | Filter „meine“ |
| Begleitung | Begleitung | text multiline | Anzeige / Legacy |
| BegleitungCodes | Begleitung-Codes | text | Lehrer-Kürzel `;`-getrennt |
| ExterneVortragende | Externe Vortragende | text | |
| Ort | Ort / Ziel | text | |
| Ablauf | Ablaufbeschreibung | multiline | |
| Unterrichtsbezug | Unterrichtsbezug | text | |
| Startdatum | Startdatum | dateOnly | |
| Enddatum | Enddatum | dateOnly | inklusiv; = Start bei eintägig |
| StartUe | Start-UE | number | Unterrichtseinheit (1–10) |
| EndUe | End-UE | number | Unterrichtseinheit (1–10) |
| StartZeit | Startzeit | text | aus UE abgeleitet bzw. Legacy `HH:mm` |
| EndZeit | Endzeit | text | aus UE abgeleitet bzw. Legacy |
| Status | Status | choice | `entwurf`, `beantragt`, `kv_stellungnahme`, `nachforderung`, `genehmigt`, `abgelehnt`, `abgesagt` |
| Notiz | Notiz / Begründung | multiline | |
| AblehnungsGrund | Ablehnungsgrund | multiline | |
| RueckfrageText | Rückfrage / Nachforderung | multiline | |
| AenderungsLog | Änderungsprotokoll | multiline | |
| BeantragtVon | Beantragt von | text UPN | |
| GenehmigtVon | Genehmigt von | text UPN | |
| GenehmigtAm | Genehmigt am | dateTime | |
| UntisErledigt | In Untis verarbeitet | boolean | Admin-Abschluss |
| UntisErledigtVon | Untis erledigt von | text | |
| UntisErledigtAm | Untis erledigt am | dateTime | |
| Verkehrsmittel | Verkehrsmittel | text | optional |
| KostenHinweis | Kostenhinweis | text | optional |
| SchulterminKey | Schultermin-Key | text | optional Sync |

### 3.2 `Aktivitaet-Regelwerk` (Seed)

| Feld | Default | Zweck |
|------|---------|--------|
| Title | Standard Schulaktivitäten | |
| RegelwerkId | `akt-rw-1` | |
| MinVorlaufTage | 7 | Antrag ≥ X Tage vor Start |
| MaxGleichzeitigProKlasse | 1 | Warnung/Fehler bei Überlappung genehmigt+beantragt |
| Aktiv | true | |

---

## 4. App-Funktionen (MVP)

| View | Lehrer | Admin |
|------|--------|-------|
| Dashboard (KPIs: offen / genehmigt / demnächst) | ja | ja |
| Liste + Filter | eigene / alle | alle |
| Kalender (Monatsraster) | ja | ja |
| Neuer Antrag / bearbeiten | eigene `beantragt` | ja |
| Genehmigen / Ablehnen | nein | ja |
| ICS-Download | gefiltert | gefiltert |
| ICS-Feed veröffentlichen / Abo-Link | ja (Site-Rechte) | ja |
| Genehmigt → Outlook / ICS / Einladen | Detail-Drawer | Detail-Drawer |
| Regelwerk anzeigen/anpassen | nein | ja |

---

## 5. Dateien

```text
docs/schulaktivitaeten-planer.md
docs/demo-data/schulaktivitaeten-2026-27.json
tools/schulaktivitaeten-planer.html
tools/sharepoint-liste-schulaktivitaeten.html
src/tools/schulaktivitaeten-planer/
  *-schema.js, *-logic.js, *-state.js, *-graph.js,
  *-export.js, *-feed.js, *-calendar.js, *-ui.js, *-js, *-css,
  *-demo-data.js, *-demo-seed.js, *-zeitraster.js
src/tools/sharepoint/sharepoint-liste-schulaktivitaeten.js
```

### Demo SJ 2026/27

- **36 Aktivitäten** (Exkursionen, Schulaktivitäten, Veranstaltungen; genehmigt / beantragt / abgelehnt)
- Buttons **Demo** / **Demo reset** im Kopf bzw. unter Regelwerk
- Lokal sofort sichtbar; optional Upsert auf SharePoint (Seed-Tag `akt-demo-2026-27`)
- Reset löscht nur Demo-Zeilen (`akt-demo-*` / Seed-Tag in Notiz)

---

## 6. Abgrenzung zum Schularbeiten-Planer

| Schularbeiten | Schulaktivitäten |
|---------------|------------------|
| Ein Tag + Dauer Min. | Start–Ende (Mehrtage möglich) |
| Fach + LBVO-Regeln | Typ + Ort + Begleitung |
| Status `fixiert` | Status `genehmigt` |
| Terminfenster-Sperren | Vorlauf + Überlappung pro Klasse |
