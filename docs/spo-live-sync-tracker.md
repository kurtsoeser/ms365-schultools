# SPO Live Sync – Tracker

**Stand:** 2026-11-18  
**Vertrag:** [spo-live-sync.md](./spo-live-sync.md) · **Skill:** `schulplaner-tool-upgrade`

Nach jedem Tool: Status hier setzen → Prompt-Block unten bleibt als Doku.

---

## So nutzt du die Datei

1. In der **Übersicht** das nächste `⬜` oder `🔍` wählen.  
2. Zum **Prompt-Block** springen (Anker in der Übersicht).  
3. Den gesamten `text`-Block kopieren → neuen Chat.  
4. Agent setzt am Ende Status auf `✅` und schreibt eine Notiz.

| Status | Bedeutung |
|--------|-----------|
| ✅ | Fertig (Pull+Push, SPO-Item, 403 kein Wipe) |
| 🔍 | Schon Graph/SoR – nur Scan, Lücken, Smoke |
| ⬜ | Offen – Sync fehlt oder unvollständig |
| ⏸ | Bewusst kein SPO-Listen-SoR |

**Fortschritt:** ✅ 14 · 🔍 3 · ⬜ 0 · ⏸ 4  

**Als Nächstes:** [Aufgaben](#02-aufgaben) oder [Teams-Antrag](#18-teams-antrag)

---

## Übersicht

| # | Tool | Familie | Status | Datum | Notiz |
|---|------|---------|:------:|-------|-------|
| 01 | [Technik-Meldungen](#01-technik-meldungen) | betrieb | ✅ | 2026-11-18 | Referenz graph+sync |
| 02 | [Aufgaben](#02-aufgaben) | betrieb | 🔍 | — | graph/sync prüfen |
| 03 | [Ausgaben](#03-ausgaben) | betrieb | ✅ | 2026-11-18 | graph+sync Inventar/Vorgänge; Typen lokal |
| 04 | [Raumplan](#04-raumplan) | betrieb | ✅ | 2026-11-18 | graph+sync; Plan lokal; Technik-`?raum=` |
| 05 | [Projektwochen](#05-projektwochen) | projekt-buchung | ✅ | 2026-11-18 | Angebote-CRUD SoR; Bookings separat; Aktion via IT/Demo-Seed |
| 06 | [Elternsprechtag-Bookings](#06-elternsprechtag-bookings) | projekt-buchung | ⏸ | 2026-11-18 | SoR = Bookings (kein SPO-Listen-CRUD) |
| 07 | [Einsammler](#07-einsammler) | ablauf | ✅ | 2026-11-18 | graph+sync; Public→Inbox/Code→Push |
| 08 | [Elternbrief](#08-elternbrief) | ablauf | ✅ | 2026-11-18 | graph+sync dual list; Public→Inbox→Push |
| 09 | [Krankmeldung](#09-krankmeldung) | ablauf | ✅ | 2026-11-18 | graph+sync; Public→lokal/Remote/Kiosk-SPO |
| 10 | [Kollegium-Umfragen](#10-kollegium-umfragen) | ablauf | ✅ | 2026-11-18 | graph+sync Zweilisten; Demo-Seed skip; Site-Kind `kollegium` |
| 11 | [sRDP-Anmeldung](#11-srdp-anmeldung) | wahl | ✅ | 2026-11-18 | SoR + Nacharbeit-Push; Sync-Meta; Demo kein Push |
| 12 | [Wahlfach-Zuteilung](#12-wahlfach-zuteilung) | wahl | ✅ | 2026-11-18 | Soft Pull+Push 4 Listen; Admin-Push nachgezogen |
| 13 | [Schulsprecherwahl](#13-schulsprecherwahl) | wahl | ⏸ | 2026-11-18 | SoR = election-api (Azure Tables); kein SPO-Listen-CRUD (Anonymität) |
| 14 | [Schulaktivitäten](#14-schulaktivitäten) | approval | ✅ | 2026-11-18 | SoR live; Pull-source in System-UI; Status-Docs = Schema |
| 15 | [Freistellung Schüler](#15-freistellung-schüler) | approval | ✅ | 2026-11-18 | Graph+CRUD SoR; Pull/Push im Entry; Demo Confirm; 403 kein Wipe |
| 16 | [Freistellung Lehrer](#16-freistellung-lehrer) | approval | ✅ | 2026-11-18 | SoR `lfr-graph`; Site-Kind `lehrerFreistellung`; Mapping-Tests; kein Delete (Approval) |
| 17 | [Schularbeiten](#17-schularbeiten) | register | ✅ | 2026-11-18 | SoR graph+CRUD; Pull/Push; Demo Confirm/local Writes; 403 kein Wipe; Sync-Meta; Mapping-Tests; Schüler-Read |
| 18 | [Teams-Antrag](#18-teams-antrag) | portal | 🔍 | — | Provisioning ≠ Liste |
| 19 | [Diplomarbeiten](#19-diplomarbeiten) | portal | 🔍 | — | |
| 20 | [Termine-Hub](#20-termine-hub) | kalender | ⏸ | — | Outlook = SoR |
| 21 | [IT Listen-Assistenten](#21-it-listen-assistenten) | it-setup | ⏸ | — | Setup only |

---

## Prompt-Blöcke (Copy & Paste)

### 01 Technik-Meldungen

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `betrieb` |
| Notiz | Referenz: `technik-meldungen-graph.js` + `technik-meldungen-sync.js` |

Kein Prompt nötig – fertig.

---

### 02 Aufgaben

| | |
|--|--|
| Status | 🔍 |
| Familie | `betrieb` |
| Docs | — |
| Hinweis | `aufgaben-graph.js` / `aufgaben-sync.js` schon vorhanden → Ist-Scan |

```text
SPO Live Sync für Aufgaben.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md
Referenz: src/tools/technik-meldungen/ und vorhandenes src/tools/aufgaben/ (graph+sync)
Wellen S0→S3. Ist-Scan zuerst (schon live?). Domäne nicht umbiegen.
Am Ende: Tracker #02 → ✅ + Datum + Notiz; Smoke Lehrer/Admin; SPO-Item sichtbar.
```

---

### 03 Ausgaben

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `betrieb` |
| Docs | [ausgaben.md](./ausgaben.md) |
| Notiz | `ausgaben-graph.js` + `ausgaben-sync.js` (Gegenstände+Vorgänge); Typen lokal; Site-Kind `ausgaben` |

Kein Prompt nötig – fertig.

---

### 04 Raumplan

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `betrieb` |
| Docs | [raumplan.md](./raumplan.md) |
| Notiz | `raumplan-graph.js` + `raumplan-sync.js`; Plan-Bild lokal; Technik-Deep-Link behalten |

```text
SPO Live Sync für Raumplan.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md
Referenz: src/tools/technik-meldungen/ (graph+sync)
Wellen S0→S3. Ist-Scan zuerst. Domäne nicht umbiegen (Räume/CSV/Plan; Technik-Deep-Link behalten).
Am Ende: Tracker #04 → ✅ + Datum + Notiz; Smoke + SPO-Item.
```

---

### 05 Projektwochen

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `projekt-buchung` |
| Docs | [projektwochen.md](./projektwochen.md) |
| Notiz | Angebote-CRUD SoR (`graph`+`sync`); Bookings = Export; Aktion Create via IT/Demo-Seed |

```text
SPO Live Sync für Projektwochen.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md + docs/projektwochen.md
Referenz: vorhandenes projektwochen-graph.js / projektwochen-sync.js (nicht Technik 1:1 kopieren)
Wellen S0→S3. Ist-Scan zuerst – CRUD vs Bookings-Export trennen. Domäne nicht umbiegen.
Am Ende: Tracker #05 → ✅ + Datum + Notiz; Smoke + SPO-Item; Bookings-Schritt unverändert.
```

---

### 06 Elternsprechtag-Bookings

| | |
|--|--|
| Status | ⏸ 2026-11-18 |
| Familie | `projekt-buchung` |
| Docs | — |
| Notiz | SoR = Microsoft Bookings (`bookingBusinesses` / Staff / Services). App ist Setup/Grundanlage per Graph; localStorage nur Prefs (`businessId`, Formular). Kein `*-schema`/`*-graph` SPO, kein Site-Kind, kein `sharepoint-liste-elternsprechtag`. Domäne nicht umbiegen – kein Fake-SPO-SoR. Analog Termine-Hub (Outlook). |

Kein Prompt nötig – bewusst kein SPO-Listen-SoR.

---

### 07 Einsammler

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `ablauf-kommunikation` |
| Docs | [einsammler.md](./einsammler.md) |
| Notiz | `einsammler-graph.js` + `einsammler-sync.js`; Public ohne Graph → Inbox/Code/Remote, Absender-Push nach Übernahme |

Kein Prompt nötig – fertig.

---

### 08 Elternbrief

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `ablauf-kommunikation` |
| Docs | [elternbrief.md](./elternbrief.md) |
| Notiz | `elternbrief-graph.js` + `elternbrief-sync.js` (Briefe + Lesebestätigungen); Public Magic-Link bleibt Inbox/Code, Merge pusht Lesestatus |

Kein Prompt nötig – fertig.

---

### 09 Krankmeldung

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `ablauf-kommunikation` |
| Docs | [krankmeldung.md](./krankmeldung.md) |
| Notiz | `krankmeldung-graph.js` + `krankmeldung-sync.js`; Admin Pull/Push SoR; Public lokal/Remote + optional Kiosk-SPO; Demo nicht pushen |

Kein Prompt nötig – fertig.

---

### 10 Kollegium-Umfragen

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `ablauf-kommunikation` |
| Docs | [kollegium-umfragen.md](./kollegium-umfragen.md) |
| Notiz | `kollegium-umfragen-graph.js` + `-sync.js` (Umfragen+Antworten); Site-Kind `kollegium`; Demo-Seeds nicht pushen; 403 kein Wipe |

Kein Prompt nötig – fertig.

---

### 11 sRDP-Anmeldung

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `wahl-portal` |
| Docs | [srdp-anmeldung.md](./srdp-anmeldung.md) |
| Notiz | Schon SoR (Pull/Create). Lücke geschlossen: Nacharbeit-PATCH, Sync-Meta System, Mapping-Tests, Demo/403 kein Wipe/Push |

```text
SPO Live Sync für sRDP-Anmeldung.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md + docs/srdp-anmeldung.md
Referenz: srdp-anmeldung-graph.js (schon SoR). Nur Lücken/Smoke.
Am Ende: Tracker #11 → ✅ + Datum + Notiz.
```

---

### 12 Wahlfach-Zuteilung

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `wahl-portal` |
| Docs | [wahlfach-zuteilung.md](./wahlfach-zuteilung.md) |
| Notiz | Soft Pull+Push für Runde/Angebot/Anmeldung/Zuteilung (`graph`+`sync`); `spoItemId`; Demo nicht pushen; Admin-Push war die Hauptlücke |

```text
SPO Live Sync für Wahlfach-Zuteilung.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md + docs/wahlfach-zuteilung.md
Referenz: wahlfach-zuteilung-graph.js. Ist-Scan + Lücken. Domäne nicht umbiegen.
Am Ende: Tracker #12 → ✅ + Datum + Notiz; Smoke Schüler/Admin.
```

---

### 13 Schulsprecherwahl

| | |
|--|--|
| Status | ⏸ 2026-11-18 |
| Familie | `wahl-portal` |
| Docs | [schulsprecherwahl.md](./schulsprecherwahl.md) |
| Notiz | SoR = Azure Tables via `election-api` (ElectionMeta / Voters / Ballots). Kein SPO-Listen-CRUD für Stimmen/Meta (Graph `Author` würde Anonymität brechen). SharePoint nur Stammdaten-Hydrate. Soft-Error kein Wipe; Demo `localDemoOnly` push’t nicht. Analog Elternsprechtag-Bookings / Termine-Hub. |

Kein Prompt nötig – bewusst kein SPO-Listen-SoR.

---

### 14 Schulaktivitäten

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `approval-planer` |
| Docs | [schulaktivitaeten-planer.md](./schulaktivitaeten-planer.md) |
| Notiz | SoR = SPO; graph CRUD + Regelwerk; Site `aktivitaeten`; Ownership; Demo nur nach Confirm; 403 behält lokal. Nachzug: `aktivitaetenLive` (graph/spo-rest/demo) in System-UI; Docs-Status inkl. entwurf/kv/abgesagt. |

Kein Prompt nötig – fertig.

---

### 15 Freistellung Schüler

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `approval-planer` |
| Docs | — (Konzept/Flow: freistellung-konzept, freistellung-flow-v2-*) |
| Notiz | SoR schon live: `freistellung-planer-graph.js` + Pull/Push in `freistellung-planer.js`; Site-Kind `freistellung`; Schema/Setup; Ownership/Actor; Demo nur mit Confirm; 403 wipe’t Demo nicht. Tests Mapping/Logic grün. Soft-Lücken: kein separates `*-sync.js`, kein Technik-Style Sync-Meta-Panel, kein `docs/freistellung.md`. |

Kein Prompt nötig – fertig.

---

### 16 Freistellung Lehrer

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `approval-planer` |
| Docs | [lehrer-freistellung-planer.md](./lehrer-freistellung-planer.md) |
| Notiz | SoR via `lfr-graph` (Pull/Push); Site-Kind `lehrerFreistellung` + Mapping-Tests; Demo nicht pushen; 403 kein Wipe; Delete bewusst weggelassen (Status-Pipeline). |

```text
SPO Live Sync für Lehrer-Freistellung-Planer.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md
Referenz: lfr-graph.js. Nur Ist-Scan/Smoke/Lücken.
Am Ende: Tracker #16 → ✅ + Datum + Notiz.
```

---

### 17 Schularbeiten

| | |
|--|--|
| Status | ✅ 2026-11-18 |
| Familie | `register-planer` |
| Docs | [schularbeiten-planer.md](./schularbeiten-planer.md) |
| Notiz | SoR graph+CRUD; Pull/Push Entry; Demo Confirm + local Writes ohne Push; 403→softError kein Wipe; Sync-Meta Admin; Mapping-Tests; Schüler-Read hart (submit/Admin). `*-sync.js` = Schultermine-Export (nicht Listen-SoR). |

Kein Prompt nötig – fertig.

---

### 18 Teams-Antrag

| | |
|--|--|
| Status | 🔍 |
| Familie | portal |
| Docs | [teams-antrag.md](./teams-antrag.md) |
| Hinweis | Graph oft = Team-Provisioning, nicht Listen-CRUD |

```text
SPO Live Sync für Teams-Antrag.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md + docs/teams-antrag.md
Ist-Scan: Anträge in SPO-Liste? Provisioning vs Listen-CRUD trennen.
Am Ende: Tracker #18 aktualisieren (✅ oder ⏸ mit Begründung).
```

---

### 19 Diplomarbeiten

| | |
|--|--|
| Status | 🔍 |
| Familie | portal |
| Docs | [diplomarbeiten.md](./diplomarbeiten.md) |

```text
SPO Live Sync für Diplomarbeiten.
Skill: schulplaner-tool-upgrade
Docs: docs/spo-live-sync.md + docs/spo-live-sync-tracker.md + docs/rollen-vertrag.md + docs/diplomarbeiten.md
Ist-Scan: Teams/Kanäle vs SPO-Listen. Nur SoR-Lücken.
Am Ende: Tracker #19 aktualisieren (✅ oder ⏸ mit Begründung).
```

---

### 20 Termine-Hub

| | |
|--|--|
| Status | ⏸ |
| Familie | kalender |
| Notiz | System of Record = Outlook-Kalender, keine klassische SPO-Liste |

Kein Prompt.

---

### 21 IT Listen-Assistenten

| | |
|--|--|
| Status | ⏸ |
| Familie | `it-setup` |
| Notiz | Setup only – gehört zur SoR-Pipeline der jeweiligen Fach-Tools |

Kein Prompt.

---

## Agent-Checkliste (jedes Tool)

1. Prompt ausführen (Wellen S0→S3 / Ist-Scan).  
2. In der **Übersicht** Status + Datum + Notiz setzen.  
3. Im **Prompt-Block** dieselbe Status-Zeile spiegeln.  
4. Fortschrittszeile oben anpassen.  
5. Chat: „Tracker aktualisiert – als Nächstes: #…“.
