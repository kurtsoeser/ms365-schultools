# Projektwochen – Umsetzungsplan

**Stand:** 2026-09-24  
**Status:** Phase 4 erledigt – Export/Druck, PA Status-Mail; MVP komplett  
**Name (UI):** Projektwochen  
**URL / mode:** `projektwochen` (Alias optional: `angebote`)  
**Zielgruppe:** HAK / berufsbildende Schulen – Projektwoche / Schlusstage mit Wahlangeboten  
**UI:** Vanilla JS/ESM, Look an `app.css` – analog Schularbeiten-Planer; **Admin-Oberfläche erstklassig** (Listen, Kalender, Filter)

Verwandt: Schularbeiten-Planer, Elternsprechtag-Bookings (`maximumAdvance` / Buchungsfenster), sRDP-Anmeldung, Stammdaten, `src/shared/ARCHITECTURE.md`.

---

## 1. Problem & Zielbild

In der **Projektwoche** entstehen viele Angebote: Zoobesuch, Kino, Workshop, Exkursion, Sportmodul usw.

1. **Lehrkräfte** reichen Angebote ein.
2. **Admin** verwaltet in der App (Listen, Kalender, Filter), freigibt und setzt **Buchungsstart** je Angebot.
3. Admin tippt **„Nach Bookings syncen“** → Dienste in Microsoft Bookings.
4. **Schüler:innen** buchen in der **Bookings-App**.
5. **Rückweg:** Termine/TN aus Bookings in der App (Listen, Suche, Filter, Belegung).

```text
Lehrer ──Antrag──► SharePoint „PW-Angebote“
                         │
Admin: Freigabe + BuchungAb + Listen/Kalender
                         │
              Button „Nach Bookings syncen“
                         ▼
              Bookings (1 Business = 1 Projektwoche)
                         │ N Services = N Angebote
                         ▼
              Schüler buchen in Bookings-App
                         │
                         ◄── Graph appointments
              TN-Listen / Belegung in der App
```

---

## 2. Architektur (fest: Hybrid)

| Schicht | Technologie | Verantwortung |
|---------|-------------|---------------|
| Anträge, Status, Buchungsstart | SharePoint-Listen | SoR |
| Admin-/Lehrer-UI | Eigene App | Übersicht, Freigabe, Sync-Button, Filter |
| Dienste & Buchung | Microsoft Bookings | 1 Business/Projektwoche, N Services |
| Teilnehmer | Graph → App (live) | TN-Listen, Belegung – kein zweites SoR |

---

## 3. Entscheidungen (fest)

| Thema | Entscheidung |
|-------|--------------|
| Name | **Projektwochen** (`?mode=projektwochen`) |
| Listen-Präfix | `PW-` (Projektwoche), z. B. `PW-Aktionen`, `PW-Angebote` |
| UI-Stack | Vanilla/ESM, Pattern Schularbeiten-Planer |
| Bookings-Business | **1 Projektwoche = 1 Business** |
| Services | **N Dienste** = freigegebene Angebote |
| Parallel | **Eine** aktive Projektwoche gleichzeitig |
| Schülerbuchung | Direkt in Microsoft Bookings |
| Sync | **Nur per Admin-Button** (nicht automatisch bei Freigabe) |
| Buchungsstart | **Admin je Angebot** einstellbar (`BuchungAb`); Default von der Aktion |
| Eltern-Einverständnis | **Nur Hinweistext** im Formular/Beschreibung – kein Workflow |
| TN aus Bookings | Ja – Kernfeature Phase 3 |
| Stammdaten | `teachers`, `classes`, optional `students` (Text-Codes) |

---

## 4. Rollen & Abläufe

### 4.1 Lehrer:in

1. Angebot anlegen → `beantragt`.
2. Eigene Anträge editieren/löschen bis Freigabe.
3. Nach Freigabe: lesen; Belegung sehen; Link ggf. an Klasse weitergeben.
4. Optional: Hinweistext „Eltern-Einverständnis empfohlen“ im Angebotstext – **kein Pflichtfeld**.

### 4.2 Admin

1. Freigabe-Queue: freigeben / ablehnen.
2. Pro Angebot (oder Batch): **`BuchungAb`** setzen – ab wann Schüler in Bookings buchen dürfen.
3. Listen, Kalender, Plan-Raster, Filter.
4. Button **„Nach Bookings syncen“** (einzeln oder alle freigegebenen) → Services anlegen/aktualisieren inkl. Scheduling/`maximumAdvance`.
5. Teilnehmer aus Bookings laden, suchen, filtern, exportieren.
6. Projektwoche anlegen/schließen; Business erzeugen.

### 4.3 Schüler:in

- Buchung nur in Bookings (App/Portal).
- Optional in unserer App: Lesen + Deep-Link (kein eigenes Buchungsformular).

---

## 5. Datenmodell

### 5.1 Stammdaten

| Entität | Nutzung |
|---------|---------|
| `teachers[]` | Antragsteller, Begleitung |
| `classes[]` | Zielklassen / Filter |
| `students[]` | TN-Join Klasse über E-Mail |

### 5.2 Liste `PW-Aktionen`

Eine Zeile pro Projektwoche. Genau eine mit `Status = offen`.

| Interner Name | Display | Typ | Bemerkung |
|---------------|---------|-----|-----------|
| Title | Name | text | „Projektwoche 2026/27“ |
| AktionId | Aktion-ID | text, unique | `pw-2027-06` |
| Startdatum / Enddatum | Zeitraum | dateOnly | Mo–Fr der Woche |
| BuchungAbDefault | Standard-Buchungsstart | dateTime | Default für neue/ohne Einzelwert |
| BookingsBusinessId | Bookings-Business-ID | text | |
| BookingsBusinessName | Anzeigename | text | |
| Status | Status | choice: `entwurf`,`offen`,`geschlossen` | |
| Beschreibung | Beschreibung | multiline | |

### 5.3 Liste `PW-Angebote`

| Interner Name | Display | Typ | Bemerkung |
|---------------|---------|-----|-----------|
| Title | Titel | (Standard) | |
| AngebotId | Angebot-ID | text, unique | |
| AktionId | Aktion-ID | text, indexed | |
| Beschreibung | Beschreibung | multiline | inkl. optionaler Eltern-Hinweis |
| HinweisEltern | Hinweis Eltern | multiline | z. B. „Einverständnis empfohlen“ – **nur Text** |
| Ort / Treffpunkt | | text | |
| Tag | Tag | choice Mo–Fr | |
| Datum | Datum | dateOnly | |
| Slot | Slot | choice: ganztags/vormittag/nachmittag/abend | |
| Startzeit / Endzeit | | text `HH:mm` | |
| Kapazitaet | Max. TN | number | → `maximumAttendeesCount` |
| PreisEuro / KostenHinweis | | number / multiline | |
| Zielklassen | | text | `1AK,1BK` oder `alle` |
| LehrerCode / LehrerEmail / Begleitung | | text | |
| Kategorie | | choice | exkursion, workshop, kultur, sport, sonstiges |
| Status | | choice | entwurf, beantragt, freigegeben, abgelehnt, abgesagt |
| **BuchungAb** | **Buchung möglich ab** | **dateTime** | **Admin; steuert Bookings `maximumAdvance`** |
| AblehnungsGrund | | multiline | |
| BookingsServiceId / BookingsBookingUrl | | text | nach Sync |
| SyncStatus / SyncFehler / SyncAm | | text / multiline / dateTime | Admin-Button-Ergebnis |
| BeantragtVon / FreigegebenVon / FreigegebenAm | Meta | | |
| NotizIntern | | multiline | nur Admin |

### 5.4 Buchungsfenster → Bookings

Vorbild: `elternsprechtag-bookings-logic.js` → `maximumAdvanceForOpenDate` + `buildSingleDayServiceSchedulingPolicy`.

| App-Feld | Bookings |
|----------|----------|
| `BuchungAb` (sonst `BuchungAbDefault` der Aktion) | `schedulingPolicy.maximumAdvance` so, dass der Termin erst ab diesem Zeitpunkt buchbar ist |
| Angebots-Datum + Start/Ende | customAvailabilities / Dauer |
| Kapazitaet | `maximumAttendeesCount` |

Effektive Öffnung: `BuchungAb` des Angebots, Fallback Aktion-Default. Admin kann vor Sync pro Karte/Tabelle ändern; nach Änderung erneut **Sync-Button**.

### 5.5 Bookings Write / Read

**Write (Sync-Button):** Business der Aktion + Services für `freigegeben`e Angebote.

**Read:** `calendarView` / `appointments` (+ Detail für `customers[]`) → Belegung, TN-Listen, Suche, CSV. Live + Button „Aktualisieren“.

---

## 6. Regeln (Logic + Vitest)

| Prüfung | Schwere |
|---------|---------|
| Pflichtfelder Titel, Datum, Kapazität ≥ 1, Lehrer | Fehler |
| Datum im Aktionsfenster | Fehler |
| Preis ≥ 0 | Fehler |
| `BuchungAb` ≤ Angebots-Datum (sinnvoll) | Warnung/Fehler |
| Sync nur bei `freigegeben` | Fehler (Guard) |
| Zielklassen / Doppelungen / Lehrer-Überschneidung | Warnung |

---

## 7. UI (Admin-first)

`tools/projektwochen.html`

| View | Inhalt | Prio |
|------|--------|------|
| `dashboard` | KPIs inkl. „Buchung offen / noch gesperrt“ | MVP |
| `kalender` | Woche/Tag, Badges Belegung + BuchungAb | MVP |
| `liste` | Filterbare Tabelle inkl. BuchungAb, SyncStatus | MVP |
| `plan` | Mo–Fr × Slots | MVP |
| `admin` | Freigabe-Queue, BuchungAb setzen, **Sync-Button** | MVP |
| `teilnehmer` | TN aus Bookings | Phase 3 |
| `neu` / `meine` | Lehrer | MVP |
| `bookings` | Business, Sync-Log, Batch-Sync | MVP |
| `setup` | Listen + Aktion anlegen | Phase 1 |
| `export` | Plan + TN | Phase 3–4 |

Filter: Tag, Slot, Status, Kategorie, Lehrer, Klasse, SyncStatus, „Buchung bereits möglich“.

---

## 8. Dateischnitt

```text
docs/projektwochen.md
tools/projektwochen.html
tools/sharepoint-liste-projektwochen.html

src/tools/projektwochen/
  projektwochen.js
  projektwochen-ui.js
  projektwochen-state.js
  projektwochen-logic.js
  projektwochen-graph.js
  projektwochen-bookings.js      ← Sync-Button Write + Appointments Read
  projektwochen-schema.js
  projektwochen-export.js

src/tools/sharepoint/sharepoint-liste-projektwochen.js

tests/projektwochen-logic.test.mjs
tests/projektwochen-bookings.test.mjs
```

| Datei | Änderung |
|-------|----------|
| `index.html` | Kachel „Projektwochen“ |
| `ms365-schooltool.html` | `?mode=projektwochen` |
| `hilfe.html` | Artikel |
| Intranet-Hub | optional Paket |
| Config | Bookings-Scopes |

**Scopes:** `Sites.ReadWrite.All` + Bookings Read/ReadWrite/Manage (wie Elternsprechtag).

---

## 9. Phasen

### Phase 0 ✅

- [x] Hybrid, 1 Business/Aktion, N Services, eine Aktion
- [x] Schüler in Bookings-App
- [x] Admin-UI prioritär, TN-Rückweg
- [x] Name **Projektwochen**
- [x] Sync **nur Admin-Button**
- [x] Eltern nur Hinweis
- [x] **BuchungAb** je Angebot (Admin)

### Phase 1 – Provisioning ✅

- [x] Schema `PW-Aktionen` / `PW-Angebote` (`projektwochen-schema.js`)
- [x] Setup-Seite `tools/sharepoint-liste-projektwochen.html`
- [x] Hub-Checkbox + `window.ms365SpoProjektwochen`
- [x] Seed-Demo-Aktion wenn PW-Aktionen leer
- [x] Dashboard-Kachel, `?mode=projektwochen`, Hilfe

### Phase 2 – Antrag + Admin-UI ✅

- [x] Formular, Freigabe, Liste/Kalender/Plan/Filter
- [x] BuchungAb in Admin-UI (+ Aktions-Default)
- [x] Demo lokal (`projektwochen-demo-data.js`)
- [x] Demo-JSON `docs/demo-data/projektwochen-demo.json` (kurtrocks) → lokal + optional SPO-Seed; Stammdaten im Browser-Backup
- [x] App `tools/projektwochen.html`, Dashboard-Kachel, Hilfe

### Phase 3 – Bookings ✅

- [x] Business je Aktion anlegen/binden (`ensureBookingBusiness`)
- [x] Sync-Button → Services + `maximumAdvance` aus BuchungAb
- [x] Appointments → Teilnehmer-View, Belegungs-Badges, CSV
- [x] Sync-Status / Fehlerlog in Admin-Bookings-Ansicht

### Phase 4 – Feinschliff ✅

- [x] Export-Ansicht: Angebote-CSV, iCal, Druck Wochenplan, TN CSV/Druck
- [x] Hilfe + Dashboard (App, Listen, PA Status-Mail)
- [x] PA-Rezept `pa-projektwochen-mail` (Freigabe/Ablehnung)

### Phase 5 – Später

- Warteliste, Entra-Rollen, TN-Archiv-Snapshot

---

## 10. Formularfelder

| Feld | MVP |
|------|-----|
| Titel, Beschreibung | ja |
| HinweisEltern (reiner Text) | ja (optional ausfüllen) |
| Datum, Slot, Start/Ende | ja |
| Ort, Treffpunkt | ja / opt. |
| Kapazität, Preis, KostenHinweis | ja |
| Zielklassen, Lehrer, Begleitung, Kategorie | ja |
| **BuchungAb** (Admin) | **ja** |
| Mitnahme / Wetter / Budget / Max. je Klasse | opt. |

---

## 11. Abgrenzung

| In Scope | Out of Scope |
|----------|--------------|
| Freigabe + Sync-Button + BuchungAb | Auto-Sync bei Freigabe |
| TN aus Graph in der App | Eigenes Buchungsformular |
| Eltern-Hinweistext | Eltern-Workflow / Unterschrift |
| 1 Business pro Projektwoche | Parallele Projektwochen |

---

## 12. Nächster Schritt

MVP (Phasen 1–4) ist fertig. Optional später Phase 5: Warteliste, Entra-Rollen, TN-Archiv.
