# Schulaktivitäten-Planer (Exkursionen & schulische Aktivitäten)

**Stand:** 2026-09-28  
**Status:** MVP in Umsetzung  
**Vorbild:** [`schularbeiten-planer.md`](./schularbeiten-planer.md)  
**Ziel:** Lehrer:innen beantragen Exkursionen / schulische Aktivitäten; Direktion/Admin genehmigt oder lehnt ab. Übersichten als Liste + Kalender, Export als ICS. **System of Record:** SharePoint-Listen auf der Intranet-Site.

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
| Statusfluss | `beantragt` → `genehmigt` \| `abgelehnt` |
| Rollen (MVP) | Demo-Umschalter Lehrer / Admin; später Entra-Gruppen |
| Backend | Kein eigenes Azure; Graph wie Schularbeiten |
| ICS | Export genehmigter (oder gefilterter) Einträge; Mehrtage als DATE-Range |
| Schultermine-Sync | optional später (Flag wie beim SA-Planer) |

**Nicht im MVP:** Elternfreigaben, Kostenabrechnung, Busbuchung, komplexes Mehrrunden-Workflow.

---

## 3. SharePoint-Listen

### 3.1 `Schulaktivitaeten`

| Interner Name | Display | Typ | Bemerkung |
|---------------|---------|-----|-----------|
| Title | Titel | Standard | z. B. Museumsbesuch |
| AktivitaetId | Aktivitäts-ID | text | `akt-` + kurz |
| Typ | Typ | choice | `Exkursion`, `Schulaktivitaet`, `Veranstaltung`, `Sonstiges` |
| KlasseCode | Klasse-Code | text | Stammdaten |
| LehrerCode | Lehrer-Kürzel | text | Antragsteller |
| LehrerEmail | Lehrer-E-Mail | text | Filter „meine“ |
| Begleitung | Begleitung | text multiline | weitere Lehrkräfte |
| Ort | Ort / Ziel | text | |
| Startdatum | Startdatum | dateOnly | |
| Enddatum | Enddatum | dateOnly | inklusiv; = Start bei eintägig |
| StartZeit | Startzeit | text | optional `HH:mm` |
| EndZeit | Endzeit | text | optional |
| Status | Status | choice | `beantragt`, `genehmigt`, `abgelehnt` |
| Notiz | Notiz / Begründung | multiline | |
| AblehnungsGrund | Ablehnungsgrund | multiline | |
| BeantragtVon | Beantragt von | text UPN | |
| GenehmigtVon | Genehmigt von | text UPN | |
| GenehmigtAm | Genehmigt am | dateTime | |
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
| ICS-Export | gefiltert | gefiltert |
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
  *-export.js, *-ui.js, *-js, *-css,
  *-demo-data.js, *-demo-seed.js
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
