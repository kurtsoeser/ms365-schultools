# Schularbeiten-Planer – Umsetzungsplan

**Stand:** 2026-09-23  
**Status:** Phase 5 erledigt – Intranet-Hinweise, Schultermine-Sync, PA Status-Mail, SA-FachMeta, IT-Checkliste
**Zielgruppe:** HAK / berufsbildende Schulen (SchUG § 17, LBVO § 7)  
**UI:** Vanilla JS/ESM, Look an `app.css` – **kein** React/Shadcn aus dem Demo-Prototyp  

Verwandt: Schultermine-Liste, Intranet-Hub, Stammdaten (`tenant-settings` / `app-data-v2`), Architektur-Leitfaden `src/shared/ARCHITECTURE.md`.

---

## 1. Zielbild

Lehrer:innen beantragen Schularbeitstermine; Admin/Direktion fixiert oder lehnt ab. Die App prüft live gesetzliche und schulautonome Regeln. **System of Record** sind SharePoint-Listen auf der Schul-Intranet-Site. Stammdaten (Klassen, Lehrer, Fächer) kommen aus dem **bestehenden Toolset**, nicht aus parallelen Listen.

```text
Lehrer / Admin ──MSAL──► Schularbeiten-Planer (ms365-schultools)
                              │ Graph (Sites.ReadWrite.All)
                              ▼
                     SharePoint Intranet-Site
                     ├─ Schularbeiten
                     ├─ Terminfenster
                     └─ Regelwerk
                              │ optional PA
                              ▼
                     Outlook / Schultermine / Mail
```

---

## 2. Entscheidungen (fest)

| Thema | Entscheidung |
|-------|--------------|
| UI-Stack | Vanilla/ESM im bestehenden Vite-MPA; Screens + Regel-Engine aus dem Prototyp **portieren** |
| Datenhaltung | SharePoint-Listen auf Intranet-Site |
| Stammdaten | **Wiederverwenden** – keine Listen `Klassen` / `Lehrer` / `Faecher` als Lookup-Ziele |
| Referenzen | Text-Codes: `FachCode`, `KlasseCode`, `LehrerCode` (+ optional E-Mail/User) |
| Rollen | Entra-Gruppen `Schularbeiten-Lehrer` / `Schularbeiten-Admin` / optional `Schularbeiten-Schüler`; **Entra Global Administrator** erhält immer Planer-Admin (Graph `RoleManagement.Read.Directory`) |
| Backend | Kein eigenes Azure für CRUD; delegiertes Graph wie bei Schulterminen |
| SPFx / Power Apps | Nicht als Primär-UI |
| Kalender-Sync | Optional: bei Status `fixiert` Eintrag in `Schultermine` (Kategorie Prüfung) oder PA-Flow |
| Fach-Meta | MVP: Defaults in Code/Regelwerk; optional später Liste `SA-FachMeta` |

---

## 3. Abgrenzung zum Demo-Prototyp

| Demo | Umsetzung hier |
|------|----------------|
| 6 Listen inkl. Lookups | 3 Listen + Stammdaten-Codes |
| React + Shadcn + TanStack Query | Vanilla + bestehendes UI-Pattern |
| Generierte `src/api/services` | `*-graph.js` über `ms365SpoGraph` |
| Rollen-Umschalter produktiv | Nur Demo; produktiv Gruppenmitgliedschaft |
| Eigenes Fächer-Farbschema in Liste | Defaults / optional `SA-FachMeta` später |

---

## 4. Datenquellen

### 4.1 Stammdaten (bereits im System)

Quelle: `ms365TenantSettingsLoad()` / `app-data-v2` (Schuljahr-Kontext).

| Entität | Felder (relevant) | Nutzung im Planer |
|---------|-------------------|-------------------|
| Fächer `subjects[]` | `code`, `name` | Dropdown Fach, Anzeige, Filter |
| Klassen `classes[]` | `code`, `name`, `year`, `headName`, `headEmail` | Dropdown Klasse, Regeln pro Klasse |
| Lehrer `teachers[]` | `code`, `name`, `email`, Fächer-Zuordnung | Dropdown, „Meine Schularbeiten“, Filter |
| Schüler `students[]` | `klasse`, `name`, `email` | Schüler-Rolle: Klasse aus Anmeldung |

Voraussetzung für sinnvollen Betrieb: Stammdaten in `tenant.html` gepflegt (wie bei Lehrerlisten-Export).

### 4.2 SharePoint-Listen (neu)

Anlegen auf der **Administrations-Site** (empfohlen) oder Intranet-URL aus Setup.

**Namenskonvention (Präfix `SAP` = Schularbeiten-Planer):**

| Liste | Zweck |
|-------|--------|
| `SAP-Regelwerk` | Max/Tag, Fristen, Sperre Notenkonferenz |
| `SAP-Terminfenster` | erlaubt/gesperrt je Zeitraum |
| `SAP-Schularbeiten` | Anträge und Termine |
| `SAP-FachMeta` | optional Farbe, Kontingent, Dauer |

Beim Setup werden **Legacy-Namen** (`Regelwerk`, `Terminfenster`, `Schularbeiten`, `SA-FachMeta`) erkannt und per Graph in die SAP-Namen **umbenannt**; fehlende Spalte `Schuljahr` wird ergänzt.

**Mehrere Schuljahre:** Einmalige Site-Einrichtung; im Planer wählt man das Schuljahr (Dropdown). Neue Einträge erhalten `Schuljahr` (z. B. `2026/27`). Alte Zeilen ohne Spalte werden für die Filterung aus dem **Datum** abgeleitet (Schuljahr ab 1. September). Klassen, Fächer und Lehrer kommen weiter aus den Stammdaten – pro Jahr können Codes wechseln, ohne die Listen neu anzulegen.

---

## 5. Listen-Schema (final für MVP)

Spalten-Definitionen im Graph-Format analog `sharepoint-liste-schultermine.js` (`text`, `number`, `boolean`, `dateTime`, `choice`).

### 5.1 Liste `SAP-Regelwerk`

| Interner Name | Display | Graph-Typ | Bemerkung |
|---------------|---------|-----------|-----------|
| Title | Name | (Standard) | z. B. „Standard HAK Regelwerk“ |
| RegelwerkId | Regelwerk-ID | text, unique/indexed | z. B. `rw-1` |
| MaxProTag | Max. pro Tag | number, default 1 | |
| MaxProWoche | Max. pro Woche | number, default 2 | |
| AnkuendigungsfristTage | Ankündigungsfrist (Tage) | number, default 7 | |
| SperreVorNotenkonferenzTage | Sperre vor Notenkonferenz | number, default 7 | |
| Aktiv | Aktiv | boolean, default true | genau ein aktiver Satz pro Schuljahr empfohlen |
| Schuljahr | Schuljahr | text | z. B. `2026/27` |

**Seed beim Setup:** ein Eintrag „Standard HAK Regelwerk“ mit Defaults und aktuellem Schuljahr.

### 5.2 Liste `SAP-Terminfenster`

| Interner Name | Display | Graph-Typ | Bemerkung |
|---------------|---------|-----------|-----------|
| Title | Titel | (Standard) | Herbstferien, Sportwoche, … |
| TerminfensterId | Terminfenster-ID | text, unique/indexed | |
| Typ | Typ | choice: `gesperrt`, `erlaubt` | default `gesperrt` |
| Startdatum | Startdatum | dateTime dateOnly | |
| Enddatum | Enddatum | dateTime dateOnly | inklusiv |
| Beschreibung | Beschreibung | text multiline | |
| Schuljahr | Schuljahr | text | Fenster gilt für dieses Schuljahr |

### 5.3 Liste `SAP-Schularbeiten`

| Interner Name | Display | Graph-Typ | Bemerkung |
|---------------|---------|-----------|-----------|
| Title | Titel (Liste) | (Standard) | Automatisch: `Schularbeit - {Fach} - {Klasse}` |
| Titel | Titel | text | Bezeichnung der Schularbeit (Import/Formular) |
| Thema | Thema | text multiline | Optional: inhaltliches Thema |
| SchularbeitId | Schularbeit-ID | text, unique/indexed | Client-generiert (`sa-` + UUID-kurz) |
| FachCode | Fach-Code | text, indexed | = `subjects[].code` |
| KlasseCode | Klasse-Code | text, indexed | = `classes[].code` |
| LehrerCode | Lehrer-Kürzel | text, indexed | = `teachers[].code` |
| LehrerEmail | Lehrer-E-Mail | text | für Filter „meine“ ohne Code-Match |
| Datum | Datum | dateTime dateOnly, indexed | |
| BeginnUhrzeit | Beginn (Uhrzeit) | text `HH:mm` | Ende wird im Planer aus **DauerMinuten** berechnet (nicht in SP gespeichert) |
| DauerMinuten | Dauer (Min.) | number, default 100 | 50–300 |
| Semester | Semester | choice: `WS`, `SS` | |
| Status | Status | choice: `beantragt`, `fixiert`, `abgelehnt` | default `beantragt` |
| Notiz | Notiz | text multiline | |
| AblehnungsGrund | Ablehnungs-Grund | text multiline | |
| BeantragtVon | Beantragt von | text (UPN) oder später User-Feld | MVP: Text UPN |
| FixiertVon | Fixiert von | text (UPN) | |
| FixiertAm | Fixiert am | dateTime | |
| Schuljahr | Schuljahr | text | Pflicht für Mehrjahresbetrieb (Import setzt mit) |

**Kein Lookup** auf andere Listen. Anzeige-Namen (Fach/Klasse/Lehrer) zur Laufzeit aus Stammdaten joinen.

### 5.4 Optional: `SAP-FachMeta`

Nur wenn Farbe / Kontingent / Standarddauer nicht in Stammdaten sollen:

- `FachCode` (unique), `Farbe`, `HatSchularbeiten`, `ProSemester`, `StandardDauer`

MVP kommt ohne diese Liste aus (Defaults: 2/Semester, 100 Min., feste Farbpalette nach Code-Hash).

### 5.5 Empfohlene SharePoint-Ansichten (manuell oder Setup)

- Offene Anträge: `Status eq beantragt`, Sort Datum  
- Fixiert nach Klasse: Filter `fixiert`, Gruppierung `KlasseCode`  
- Aktive Sperrzeiten: `Enddatum >= heute`

---

## 6. Regel-Engine (Port aus Prototyp)

Reine Logik in `schularbeiten-planer-logic.js` – **ohne DOM/fetch**, Vitest.

### 6.1 Eingaben

- Entwurf einer Schularbeit (Codes, Datum, Dauer, Semester, Status)  
- Bestehende Schularbeiten derselben Klasse (Status `beantragt` + `fixiert`)  
- Aktives Regelwerk  
- Terminfenster  
- Optional: Fach-Meta / Defaults  

### 6.2 Prüfungen

| Regel | Schwere | Grundlage |
|-------|---------|-----------|
| Max. N pro Tag pro Klasse | Fehler | LBVO § 7 Abs. 8 |
| Max. M pro Kalenderwoche pro Klasse | Fehler | LBVO § 7 Abs. 8 |
| Ankündigungsfrist ≥ X Tage | Fehler | LBVO § 7 Abs. 1 |
| Datum in gesperrtem Terminfenster | Fehler | schulautonom |
| Sperre vor Notenkonferenz (über Terminfenster oder Tage-Regel) | Fehler | schulautonom |
| Tag nach schulfreien Tagen | Warnung | LBVO § 7 Abs. 8 |
| Kontingent pro Fach/Klasse/Semester | Warnung oder Fehler (konfigurierbar; MVP: Warnung) | Lehrplan |
| Dauer außerhalb 50–150 (bzw. Meta) | Warnung | Lehrplan |

Ausgabe: `{ errors: string[], warnings: string[] }`. Speichern nur wenn `errors.length === 0` (außer Admin-Override – **nicht** im MVP).

---

## 7. Rollen & Berechtigungen

### 7.1 App-Logik

| Aktion | Lehrer | Admin | Schüler |
|--------|--------|-------|---------|
| Antrag erstellen | ja | ja | nein |
| Eigene `beantragt` bearbeiten/löschen | ja | ja | nein |
| Alle sehen / filtern | eingeschränkt (eigene) | ja | nein |
| Nur fixierte der eigenen Klasse | – | – | ja |
| Fixieren / Ablehnen / Verschieben | nein | ja | nein |
| Terminfenster / Regelwerk | nein | ja | nein |
| Export | ja (gefiltert) | ja | ja (Klasse, nur fixiert) |

**Schüler-Zuordnung:** Anmeldung (E-Mail) → Eintrag in Stammdaten `students[]` (`klasse`, `name`, `email`). Ohne Treffer (Demo): Klassen-Auswahl in der Sidebar bzw. `?role=schueler&klasse=3AK`.

Rollenauflösung MVP:

1. Graph: Mitgliedschaft in konfigurierbaren Gruppen-IDs/Namen (Config oder Stammdaten-Feld)  
2. Fallback Dev: UI-Umschalter hinter `?demoRole=1` oder localStorage-Flag  

### 7.2 SharePoint (Empfehlung an die Schule)

**Site:** Planungslisten auf einer **Administrations-Site** (z. B. `/sites/administration`), nicht auf dem öffentlichen Intranet.

**Provisioning:** Tool „Schularbeiten-Listen“ oder Planer → Administration → **SharePoint-Berechtigungen** setzen Entra-Gruppen und brechen die Vererbung pro Liste (breite Site-Besucher/Mitglieder werden entfernt). Voraussetzung für die ausführende Person: `Sites.FullControl.All`.

| Liste | Schulleitung (Admin) | Lehrer-Gruppe | Schüler-Gruppe |
|-------|------------|---------------|----------------|
| **Schularbeiten** | Vollzugriff | Beitragen | Lesen |
| Regelwerk, Terminfenster, SA-FachMeta | Vollzugriff | Lesen | *(kein Zugriff)* |

Gruppennamen (Anzeigename, Mail oder Objekt-ID) werden in `localStorage` (`ms365-schularbeiten-perms-v1`) geteilt zwischen Setup-Tool und Planer.

Durchsetzung „nur eigene bis beantragt“ und „Schüler nur fixiert + Klasse“ zusätzlich in der App (SharePoint „nur eigene Elemente“ optional manuell in den Listen-Einstellungen).

---

## 8. Datei- und Tool-Schnitt

### 8.1 Neue Dateien

```text
tools/schularbeiten-planer.html          ← Planer-UI (Hauptwerkzeug)
tools/sharepoint-liste-schularbeiten.html ← Setup: Listen anlegen (optional eigene Seite
                                             oder Tab im Planer „Einrichtung“)

src/tools/schularbeiten-planer/
  schularbeiten-planer.js                ← Entry / Wiring
  schularbeiten-planer-ui.js             ← Render (Sidebar-Views, Formulare, Kalender)
  schularbeiten-planer-state.js          ← Filter, Rolle, aktuelle Ansicht
  schularbeiten-planer-graph.js          ← CRUD Listen via ms365SpoGraph
  schularbeiten-planer-logic.js          ← Regel-Engine + Hilfen (testbar)
  schularbeiten-planer-export.js         ← iCal / Print-Daten
  schularbeiten-planer-schema.js         ← Spaltendefinitionen für Setup

src/tools/schularbeiten-planer/
  schularbeiten-planer-logic.test.mjs    ← Vitest
```

Alternative: Setup-Logik in `src/tools/sharepoint/sharepoint-liste-schularbeiten.js` spiegeln (wie Schultermine) und vom Hub aufrufen – empfohlen für Konsistenz.

### 8.2 Bestehende Dateien anpassen

| Datei | Änderung |
|-------|----------|
| `index.html` | Kachel(n): Setup + Planer, Cluster `website` (und/oder `unterricht`) |
| `ms365-schooltool.html` | `?mode=schularbeiten` → Planer |
| `sharepoint-intranet-hub.js` | Option „Schularbeiten-Paket“ anlegen |
| `hilfe.html` | Abschnitt Tool-Hilfe |
| `pa-recipes-catalog.js` | optional Rezept Status-Mail / Sync |

### 8.3 Scopes

Wie Schultermine: `User.Read`, `Sites.ReadWrite.All`. Für Rollengruppen optional `GroupMember.Read.All` (oder Gruppen-IDs in Config und `/me/memberOf`).

---

## 9. UI-Struktur (Vanilla, ein HTML)

Eine Seite mit **interner Navigation** (Sidebar wie Prototyp, Styles aus `app.css`):

| View-ID | Inhalt | Priorität |
|---------|--------|-----------|
| `dashboard` | KPI-Karten, optional einfache Wochenverteilung (Canvas/SVG, kein d3-Zwang) | MVP-spät |
| `kalender` | Monat (+ optional Woche) | MVP |
| `neu` | Antragsformular + Live-Regelprüfung | MVP |
| `meine` | Tabelle eigener Anträge | MVP |
| `admin` | Offene Anträge, Terminfenster, Regelwerk | MVP |
| `regeln` | Info-Texte SchUG/LBVO | MVP (statisch) |
| `export` | iCal, Druck/PDF | MVP |
| `setup` | Site-URL, Listen anlegen, Seed Regelwerk | Phase 1 |

Globale Filterleiste: Klasse, Fach, Lehrer, Status (wirkt auf Kalender/Listen/Export).

---

## 10. Phasen & Tickets

### Phase 0 – Spezifikation (dieses Dokument)

- [x] Architektur-Entscheidung Vanilla + 3 Listen  
- [ ] Review Schema mit Pilotschule (Gruppennamen Entra)  
- [ ] Intranet-Site-URL-Konvention festhalten  

**Fertig wenn:** Schema und Rollen von dir abgenommen.

---

### Phase 1 – Provisioning (Setup-Tool)

**T1.1** ✅ Spaltendefinitionen in `schularbeiten-planer-schema.js` (Graph-JSON)  
**T1.2** ✅ `createLists(webUrl)` – drei Listen idempotent (finden oder anlegen + fehlende Spalten)  
**T1.3** ✅ Seed: aktives Standard-Regelwerk  
**T1.4** ✅ HTML-Setup-UI (`tools/sharepoint-liste-schularbeiten.html`)  
**T1.5** ✅ Hook Intranet-Hub-Startpaket (Checkbox)  
**T1.6** ✅ Dashboard-Kachel + Hilfe + `?mode=schularbeiten`  

**Fertig wenn:** Auf Test-Site erscheinen drei Listen; Regelwerk hat einen aktiven Eintrag.

---

### Phase 2 – Kern-Logik

**T2.1** ✅ Regel-Engine (`schularbeiten-planer-logic.js`)  
**T2.2** ✅ Vitest: Tag/Woche/Frist/Sperrfenster/Warnung nach Sperrzeit  
**T2.3** ✅ Graph-CRUD: Items lesen/schreiben/patchen/löschen  
**T2.4** ✅ Join Stammdaten ↔ Codes (Label-Resolver)  
**T2.5** ✅ Rollenauflösung (Entra-Gruppen aus Administration; Demo mit `?demoRole=1`)  

**Fertig wenn:** Tests grün; manuell Item in Liste anlegbar über Graph-Hilfsfunktion.

---

### Phase 3 – MVP-UI

**T3.1** ✅ Shell: Layout, Nav, Filter, Skeleton-Zustände  
**T3.2** ✅ View `neu` – Formular + Live-Prüfung + Speichern `beantragt`  
**T3.3** ✅ View `meine` – Tabelle, Bearbeiten/Löschen wenn `beantragt`  
**T3.4** ✅ View `admin` – Fixieren/Ablehnen (+ Ablehnungsgrund)  
**T3.5** ✅ View `admin` – Terminfenster CRUD  
**T3.6** ✅ View `admin` – Regelwerk laden/speichern  
**T3.7** ✅ View `kalender` – Monat, Statusfarben, Klick → Detail  
**T3.8** ✅ View `regeln` – statische Info-Karten  
**T3.9** ✅ View `export` – .ics + Print-CSS  
**T3.10** ✅ Dashboard-Kachel Planer + `mode=` + Hilfe-Abschnitt  

**Fertig wenn:** Lehrer-Antrag → Admin fixiert → erscheint im Kalender/Export; Regelverletzung blockiert Speichern.

---

### Phase 4 – Dashboard & Feinschliff

**T4.1** ✅ View `dashboard` – KPIs inkl. Konflikte  
**T4.2** ✅ Einfache Wochenverteilung (SVG, gestapelt nach Fach)  
**T4.3** ✅ Gesperrte Tage im Kalender schraffieren  
**T4.4** ✅ Detail-Dialog mit rollenabhängigen Aktionen (+ Regelhinweise)  

**Fertig wenn:** Parität zu den wichtigsten Prototyp-Screens ohne React.

---

### Phase 5 – Intranet & Automatisierung (nach MVP)

**T5.1** ✅ Navigationslink / Embed-Hinweis (Dashboard + Admin, URL/Snippet kopieren)  
**T5.2** ✅ Bei Fixierung optional Eintrag in `Schultermine` (Kategorie Prüfung, Marker `[SA:…]`)  
**T5.3** ✅ PA-Rezept „Schularbeiten Status-Mail“ (`pa-schularbeiten-mail.html`)  
**T5.4** ✅ Liste `SA-FachMeta` + Admin-UI (Farbe, Kontingent, Dauer)  
**T5.5** ✅ IT-Checkliste (unten Abschnitt 16)  

---

## 11. Ausroll bei der Schule (Betriebsablauf)

1. Stammdaten in Schultools pflegen (Klassen, Lehrer, Fächer).  
2. Entra-Gruppen anlegen und Mitglieder zuweisen.  
3. Intranet-Site vorhanden (Hub-Tool).  
4. Setup-Tool: Listen anlegen.  
5. Terminfenster (Ferien, Sportwoche, Sperrwoche) eintragen.  
6. Regelwerk prüfen/anpassen.  
7. Planer-Link ins Intranet (Quicklink) + kurze Lehrkräfte-Schulung.  
8. Pilot 1 Semester; parallel alte Excel-Liste nur als Fallback.

---

## 12. Nicht-Ziele (MVP)

- React-/Shadcn-Port 1:1  
- SharePoint-Lookups auf Klassen/Lehrer/Fächer  
- Eigenes App-Backend  
- SPFx-Webpart  
- Admin-Override bei Regelverstößen  
- Automatischer Import aus WebUntis/Stundenplan  
- ~~Entra-Gruppenauflösung produktiv~~ (umgesetzt; Dev: `?demoRole=1` für Demo-Umschalter)  

---

## 13. Risiken & Mitigation

| Risiko | Mitigation |
|--------|------------|
| Stammdaten unvollständig | Setup prüft und warnt; leere Dropdowns erklären |
| Code-Umbenennung Klasse/Lehrer | Codes stabil halten; Anzeige über Stammdaten |
| Graph-Throttling | Batch klein, Sleep wie Schultermine; Retry in `spo-graph-shared` |
| Lehrer sieht fremde Anträge | Filter in App + SP-Elementrechte |
| Doppeldefinition Prüfung vs. Schularbeit | Getrennte Listen; optional Sync nur bei `fixiert` |

---

## 14. Abnahmekriterien (MVP gesamt)

1. Listen per Tool auf Intranet-Site anlegbar.  
2. Antrag mit Live-Regelprüfung speicherbar.  
3. Admin kann fixieren/ablehnen.  
4. Kalender zeigt Termine farbig nach Status.  
5. Export .ics und Druck funktionieren mit Filtern.  
6. Keine neuen npm-Dependencies.  
7. Regel-Engine durch Vitest abgedeckt.  
8. Keine parallelen Klassen-/Lehrer-/Fächer-Listen.

---

## 15. Nächster konkreter Implementierungsschritt

MVP inkl. Phase 5 und Schüler-Ansicht (Demo) ist im Repo umgesetzt. Optional später:

- Konflikt-KPI-Drilldown  
- Automatische Wochenverteilung als Export-Grafik  

---

## 16. Demo-Daten Schuljahr 2026/27 (kurtrocks)

Umfassendes Seed-Paket für die Test-Site  
`https://kurtrocks.sharepoint.com/sites/MS365-Schultools`:

| Inhalt | Anzahl (ca.) |
|--------|----------------|
| Fächer-Meta (`SA-FachMeta`) | 12 |
| Terminfenster (Ferien, Sportwoche, Sperrwochen) | 10 |
| Schularbeiten (WS+SS, fixiert/beantragt/abgelehnt) | 59 |
| Stammdaten-Codes (Klassen/Lehrer/Fächer) | 10 / 8 / 12 |

**JSON:** [`docs/demo-data/schularbeiten-2026-27.json`](./demo-data/schularbeiten-2026-27.json)  
**Neu erzeugen:** `node scripts/generate-schularbeiten-demo-json.mjs`

### Laden auf SharePoint

1. **Schularbeiten-Planer** öffnen → **JSON importieren**  
   Datei: `docs/demo-data/schularbeiten-2026-27.json`  
   → Daten erscheinen sofort lokal; optional „Auf SharePoint schreiben“.  
2. Oder Tool **Schularbeiten-Listen** → Site-URL → **Paket anlegen** → **Demo 2026/27 laden**.  

Stammdaten (Klassen/Lehrer/Fächer) werden beim JSON-Import nach Möglichkeit in die lokalen Schul-Grundeinstellungen übernommen.

---

## 17. IT-Checkliste Schulausroll

### Einmalig (Schul-IT)

1. **Stammdaten** in Schultools pflegen (Klassen, Lehrer mit E-Mail, Fächer).  
2. **Entra-Gruppen** anlegen (Empfehlung): `Schularbeiten-Lehrer`, `Schularbeiten-Admin` – Mitglieder zuweisen.  
3. **Intranet-Site** vorhanden (Hub-Tool); Schreibrechte für Admins auf der Site.  
4. **Listen-Paket** ausführen: Regelwerk, Terminfenster, Schularbeiten, SA-FachMeta.  
5. SharePoint-Berechtigungen (Empfehlung):  
   - Regelwerk / Terminfenster / SA-FachMeta: Lehrer Lesen, Admin Bearbeiten  
   - Schularbeiten: Lehrer Beitragen; Elementbearbeitung „nur eigene“; Admin Vollzugriff  
6. Optional: **Schultermine-Liste** + Flow Termine→Kalender.  
7. Optional: Flow **Status-Mail** (`pa-schularbeiten-mail`).  
8. Planer-**Quicklink** im Intranet (URL/Snippet aus dem Planer kopieren).  

### Pro Schuljahr

1. Terminfenster (Ferien, Sportwoche, Sperrwoche) aktualisieren.  
2. Regelwerk prüfen.  
3. Fach-Meta (Farben/Kontingente) nach Bedarf.  
4. Kurzschulung Lehrkräfte (Antrag → Fixierung).  

### Betrieb

| Thema | Hinweis |
|-------|---------|
| Demo-Rolle | Nur Vorschau; produktiv Gruppen/Rechte in SharePoint + später Entra in der App |
| Sync Schultermine | Admin → Automatisierung → Checkbox; gleiche Site wie Planer |
| Doppelte Prüfungen | Schularbeiten = Workflow; Schultermine = Kalender – Sync nur bei Fixierung |
| Graph-Scopes | `Sites.ReadWrite.All` + Anmeldung |

---

## Anhang A – Mapping Demo → Umsetzung

| Demo-Feld / Liste | Umsetzung |
|-------------------|-----------|
| Faecher.* | Stammdaten `subjects` (+ optional SA-FachMeta) |
| Klassen.* | Stammdaten `classes` |
| Lehrer.* | Stammdaten `teachers` (+ Entra-Rolle) |
| Terminfenster | Liste `Terminfenster` |
| Regelwerk | Liste `Regelwerk` |
| Schularbeiten.FachLookup | `FachCode` |
| Schularbeiten.KlasseLookup | `KlasseCode` |
| Schularbeiten.LehrerLookup | `LehrerCode` (+ `LehrerEmail`) |
| BeantragtVon (User) | MVP: Text UPN |

## Anhang B – Default-Farbpalette (ohne Meta-Liste)

Deterministische Farbe aus `FachCode` (Hash → feste Palette in CSS-Variablen / Array), damit Kalender ohne SharePoint-Farbfeld funktioniert.
