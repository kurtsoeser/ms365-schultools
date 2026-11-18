# Lehrer-Portal & Schulalltag – Roadmap

**Stand:** 2026-11-08  
**Status:** Planung (Referenz-Screenshots Schul-Dashboard + Abgleich mit Bestand)  
**Kontext:** Ergänzung zu [`schul-apps-roadmap-prozesse.md`](./schul-apps-roadmap-prozesse.md), [`license-modules.md`](./license-modules.md), [`analyse-04-ausbau-neue-tools.md`](./analyse-04-ausbau-neue-tools.md)  
**Nicht im Scope:** Zahlungen/Payment, volles Absenzenbuch (bleibt Untis), Noten/LMS, zweites SIS

---

## 0. Zielbild in einem Satz

Lehrkräfte und Schulleitung bekommen ein **Alltags-Portal** (linke Sidebar, Schnellzugriff, später personalisierbares Canvas) – während IT den bestehenden Werkzeugkasten behält. Fachdaten liegen in der **SharePoint-/Graph-Welt der Schule**; Layout-Persönliches beim **User** (Browser-Cache + OneDrive), nicht auf Betreiber-Servern. Neue Apps: SharePoint SoR, schlanke UI, optional Graph/PA, Lizenz über Pakete.

**Rollenvertrag:** Portal vs. IT-Backend, Personas und App-Rollen – siehe [`rollen-vertrag.md`](./rollen-vertrag.md).

---

## 1. Ausgangslage

### 1.1 Was die Referenz-Screenshots gut machen

| Muster | Warum übernehmen |
|--------|------------------|
| Linke Sidebar nach **Schulalltag** (Termine, Veranstaltungen, Technik, …) | End-User denkt nicht in „01 Mitgliedschaften“ |
| **Schnellzugriff** mit Favoriten + externen Links (Mail, Kalender, Untis, Sokrates) | Portal statt nur Tool-Katalog |
| Rechte **Aktions-Spalte** („Neu …“, „Meine …“) | Weniger Klicks, klare Einstiege |
| **Typisierte Formulare** statt generischer Felder | ARGE-Tag ≠ Bildungsberatung ≠ Technik-Ticket |
| Große Formulare mit Beispiel-Platzhaltern | Weniger Schulung |
| Farbige Icons, ruhiges Layout | Orientierung ohne Chaos |

### 1.2 Was bei uns schon stark ist

| Thema | Bestand | Bemerkung |
|-------|---------|-----------|
| Schüler-Freistellung | `freistellung-planer` | Antrag + Freigabe |
| Lehrer-Abwesenheit | `lehrer-freistellung-planer` | inkl. iCal/Direktion |
| Schulaktivitäten | `schulaktivitaeten-planer` | Genehmigung, Kalender, ICS |
| Schularbeiten | `schularbeiten-planer` | bereits Lehrer-/Schüler-sichtbar |
| Eltern & Ablauf | Einsammler, Elternbrief, Krankmeldung, Kollegium | oft nur IT im Dashboard sichtbar |
| Ausgaben | `ausgaben` | Schlüssel/Karten – verwandt mit Technik/Raum |
| Termine (Daten) | SharePoint Schultermine + PA Sync | fehlt: Lehrer-UI + Graph→Outlook |
| Räume (IT) | `raeume-ressourcen` | schwach; kein Raum**plan**-Werkzeug |
| Dashboard | Sidebar + Favoriten | IT-Cluster; Lehrer-Whitelist sehr eng |

### 1.3 Kern-Lücke

Viele Apps existieren schon – aber **Lehrkräfte sehen sie kaum**. Die Roadmap hat zwei Schienen:

1. **Portal-Schiene** – Sichtbarkeit, Sidebar, Schnellzugriff  
2. **Produkt-Schiene** – neue/ausbauende Module (Aufgaben, Technik, Termine, Raumplan, Veranstaltungen, …)

---

## 2. UX: personalisiertes Canvas-Dashboard

### 2.1 Vision

Neben der festen Sidebar bekommt jede Person (Lehrkraft, Schüler:in, später Direktion) ein **anordenbares Dashboard**:

| Fähigkeit | Beschreibung |
|-----------|--------------|
| **Kacheln platzieren** | Drag & Drop auf einem Raster/Canvas |
| **Größe ändern** | z. B. 1×1, 2×1, 2×2 (kleine Shortcut-Kachel vs. Widget) |
| **Favorisieren** | Stern → erscheint auf dem Canvas / Schnellzugriff |
| **Ausblenden** | Katalog-Tools, die mich nicht interessieren, vom Home nehmen |
| **Widgets** | nicht nur Links: „nächste Termine“, „meine offenen Aufgaben“, „Technik offen“ |
| **Profil-Default** | Schüler vs. Lehrer starten mit sinnvollen Vorlagen, dann personalisieren |

Das muss **nicht in P0** sein. Sinnvolle Stufung:

| Stufe | Lieferobjekt | Phase |
|-------|--------------|-------|
| A | Favoriten + Reihenfolge (habt ihr schon teilweise) + Tenant-Links | P0 |
| B | Edit-Modus: Kacheln an-/aus, einfache Reihenfolge | nach P0 |
| C | **Canvas-Layout** (Position + Größe) + Widget-Kacheln | eigene UX-Phase, parallel/nach P1 |
| D | Rollen-Vorlagen („Lehrer-Standard“, „Schüler-Standard“) vom Admin vorschlagen, User überschreibt | mit C oder danach |

### 2.2 Speicherort – Multi-School-Prinzip (entscheidend)

Die App läuft zentral (Browser unter `app.ms365.schule`), **Schuldaten** liegen im jeweiligen Tenant (SharePoint/Graph). Personalisierung darf **nicht** auf Betreiber-Servern als Nutzerprofil-DB landen.

| Option | Wo liegt das Layout? | Gerätewechsel | Passt zu Multi-School? | Empfehlung |
|--------|----------------------|---------------|------------------------|------------|
| **A – localStorage** | Nur dieser Browser | nein | ja (kein Server) | MVP / Offline-Cache |
| **B – OneDrive App-Folder** | Graph `drive/special/approot` des Users | ja (gleicher M365-Account) | **ja – Ideal** | Ziel für Canvas |
| **C – SharePoint-Liste pro Schule** | Liste `UserDashboardLayouts`, Item je `UserId`/`UPN` | ja | ja (Daten in der Schule) | Alternative, wenn Admin Layouts sehen/resetten soll |
| **D – Betreiber-Backend** | eure DB | ja | **nein** – Daten bei euch | bewusst vermeiden |

**Empfohlenes Modell (Hybrid):**

```text
Browser (localStorage)
    ↕ schnelle UX, Offline
User-Graph (OneDrive approot)  ← „gehört dem User“
    optional Spiegel
Schul-SharePoint (nur wenn Schule Templates/Reset braucht)
```

Payload klein und schulbezogen (kein Fachdaten-Leak):

```json
{
  "v": 1,
  "tenantId": "<entra-tenant>",
  "persona": "lehrer",
  "updatedAt": "2026-11-08T23:00:00Z",
  "tiles": [
    { "id": "termine-hub", "x": 0, "y": 0, "w": 2, "h": 1 },
    { "id": "widget-naechste-termine", "x": 2, "y": 0, "w": 2, "h": 2 },
    { "id": "link-webuntis", "x": 0, "y": 1, "w": 1, "h": 1 }
  ],
  "hiddenToolIds": [],
  "favorites": ["schulaktivitaeten-planer", "elternbrief"]
}
```

Hinweise:

- Key immer **pro Tenant** speichern (`…-by-tenant`), weil User in mehreren Schulen vorkommen können.
- Heute: Favoriten/Order schon in `localStorage` (`ms365-dashboard-favorites-v1`, `ms365-dashboard-order-*`); SPO-Config-Bundle sync’t eher **Schul-/IT-Konfig**, nicht „Michelle’s Canvas“.
- Canvas-Layouts also **nicht** in den IT-Config-Bundle (`dashboard-access.json`) mischen – sonst landen persönliche Layouts im Schul-Backup-Monolithen.

### 2.3 UX-Leitplanken

- Edit-Modus klar getrennt („Anordnen“ / „Fertig“) – Alltag bleibt ruhig.
- Mobile: Canvas vereinfachen (nur Reihenfolge, keine freien Größen) oder feste Stack-Ansicht.
- Accessibility: Tastatur-Verschieben, nicht nur Drag.
- Keine kritischen Fachdaten im Layout-JSON (nur Tool-/Widget-IDs und Positionen).

---

## 3. Ziel-Sidebar (Lehrer / Direktion)

Vorschlag Menüstruktur (End-User). IT behält den vollen Katalog.

| Menü | Inhalt (Kacheln / Einstiege) |
|------|------------------------------|
| **Dashboard** | Begrüßung, Schnellzugriff, Favoriten, Tenant-Links |
| **Termine** | Nächste 14 Tage, Termin anlegen, Abwesenheit, Jahreskalender |
| **Schulaktivitäten** | Anmelden, Meine, Limits, Klassen-Auslastung |
| **Unterricht** | Schularbeiten, Freistellung Schüler, Diplom, Projektwochen |
| **Aufgaben** | Aufgaben an Personen / Rollen zuweisen & abhaken *(neu)* |
| **Technik** | Fehlermeldung, offene Tickets, Zugangsinfos *(neu)* |
| **Eltern & Ablauf** | Elternbrief, Einsammler, Krankmeldung |
| **Kollegium** | Umfragen, Lehrer-Freistellung |
| **Räume** | Raumplan, Belegung, Raumliste *(neu/Ausbau)* |
| **Hilfe / Intranet** | Q-Handbuch / Prozesswiki, Intranet-Hub |

Optional später: Klassengeld, Unterrichtsmaterial, Öffentlichkeitsarbeit (siehe Phase P5+).

---

## 4. Lieferphasen (Überblick)

| Phase | Fokus | Warum zuerst | Aufwand |
|-------|-------|--------------|---------|
| **P0** | Lehrer-Portal-Grundlage | **umgesetzt** | niedrig |
| **P1** | Kalender und Termine MVP (local + Outlook vormerken + Abwesenheit-Link) | **umgesetzt** (Graph-Write folgt) | mittel |
| **P2** | Schulaktivitäten: Hub-Kacheln, Kostenlimits, Auslastung, Lehrer-Vorschlag | **umgesetzt** | mittel |
| **P3** | Aufgaben MVP | **umgesetzt** | mittel |
| **P4** | Technik-Meldungen MVP | **umgesetzt** | mittel |
| **P5** | Raumplan MVP (Import + Viewer) | **umgesetzt** | mittel |
| **PX** | Canvas-Dashboard + OneDrive Sync | offen | mittel |
| **P0.1** | Portal-Katalog erweitern (Matura, Kursteams, …) + IT-Matrix greift | **umgesetzt** | niedrig |
| **P0.2** | IT-Portal-Standard (anheften, Schul-Defaults Lehrkraft/Schüler) | **umgesetzt** (Basis) | mittel |
| **P6+** | Q-Handbuch, Unterrichtsmaterial, … | offen | variabel |

Abhängigkeiten: P0 vor allem unabhängig. P1 nutzt bestehende Schultermine-Liste. P3/P4 teilen Ticket-Muster. P5 kann parallel zu P3/P4 starten. **PX** baut auf P0-Favoriten auf und wird wertvoller, sobald Termine/Aufgaben-Widgets existieren (nach P1/P3). **P0.2** baut auf Werkzeug-Zugriff-Matrix und portal-fähigem Katalog auf.

---

## 5. Phase P0 – Lehrer-Portal-Grundlage

**Status:** umgesetzt (2026-11-09) – Whitelist, gruppiertes Portal, Schnellzugriff-Links  
**Shell-Update (2026-11-09):** linke Portal-Nav + 3-Spalten-Grid; **Schulportal zuerst**, IT-Bereich bewusst  
**Katalog-Update (2026-11-18):** Portal-fähige Alltags-Apps erweitert (u. a. Maturaanmeldung/sRDP, Kursteams, Wahlfach, Schulwahlen, Intranet); Whitelist = portal-fähig, Audience/IT-Matrix entscheidet Sichtbarkeit

### Ziel

Lehrkräfte öffnen die App und sehen **Alltagstools**, nicht den IT-Werkzeugkasten. Schul-IT startet ebenfalls im Schulportal und öffnet den Vollkatalog nur über „IT-Bereich“.

### Lieferobjekte

1. **Whitelist erweitern** (`DASHBOARD_LEHRER_SCHUELER_APPS` bzw. Audience-Regeln):
   - `schulaktivitaeten-planer`
   - `lehrer-freistellung-planer`
   - `elternbrief`, `einsammler`, `krankmeldung`
   - `kollegium-umfragen`
   - bestehende: Schularbeiten, Freistellung, Teams-Antrag, Diplom
   - **nachgezogen:** `srdp-anmeldung` (Maturaanmeldung), Kursteams/OneNote, Wahlfach, Schulwahlen, Intranet, Elternsprechtag, KlassenChats, Ausgaben
2. **Portal-Sektionen** (`DASHBOARD_PORTAL_SECTIONS_BY_VIEW`): Unterricht / Schulalltag / Betrieb & Technik / Eltern & Ablauf / Kollegium / Intranet (ohne IT-Nummern).
3. **Schnellzugriff**: Outlook/Teams/Kalender + konfigurierbare Schul-Links (WebUntis, Sokrates, Intranet) – `dashboard-portal-links.js`, Speicherung `setup.portalQuickLinks` + localStorage.
4. Rollen-Hinweis im Portal-Hero („Lehrkraft“ / „Schüler/in“).
5. **Portal-Shell**: linke Nav (Dashboard + Sektionen), Hauptbereich 3-spaltig; IT-CTA nur für `isIt`.
6. **Zwei Welten, eine App**: `data-dash-app-only` (Portal) vs. `data-dash-it-shell` (IT); Auth-Menü Portal | IT | Schüler-Vorschau.

### Akzeptanz

- Demo-Lehrer sieht die lizenzierten Alltags-Apps in den Portal-Sektionen, nicht den IT-Werkzeugkasten (~90 Setup-Kacheln).
- Schul-Links über „Links bearbeiten“ setzbar; Favoriten bleiben wie bisher.
- Lehrer/Schüler sehen nie Organigramm/Stammdaten im Header (auch kein Flash).
- Schul-IT landet im Portal; IT-Bereich nur über Kachel/Nav/Auth-Menü; Zurück über Auth-Menü „Portal“.

### Manuelle Checks

| Rolle | Erwartung |
|-------|-----------|
| Lehrkraft | Portal-Sidebar, 3-Spalten-Apps, kein IT-Chrome |
| Schüler/in | „Meine Apps“, kein Links-bearbeiten, kein IT |
| Schul-IT | Start Portal + IT-Kachel; nach „IT-Bereich“ Vollkatalog; zurück über Auth-Menü „Portal“ |
| Tool-Seite (z. B. Freistellung) | Zurück zur Startseite → gleiche Shell wie zuvor (Portal bzw. IT laut Speicher) |

### Lizenz

Kein neues Paket – Sichtbarkeit bestehender Features. Tenant-Pakete filtern Kacheln für alle Rollen gleich.

### P0.2 – IT-Portal-Standard (Basis umgesetzt)

Der IT-Leiter legt pro Schule fest, welche portal-fähigen Apps Lehrkräfte/Schüler **sehen** (Matrix) und welche auf dem Home **angeheftet** sind („Portal gestalten“ auf `dashboard-werkzeug-zugriff.html`).

| Schicht | Wer | Was |
|---------|-----|-----|
| Lizenz | Verkauf / Betreiber | Welche Module die Schule darf |
| Portal-fähig | Produktkatalog | Alltags-Apps (nicht IT-Setup) |
| Schul-Standard | IT-Leiter | Sichtbarkeit + `ms365-dashboard-portal-defaults-v1` (Pins, Sidebar-Reihenfolge) |
| Persönlich | User | Favoriten (`ms365-dashboard-favorites-v1`) nach den Schul-Pins; Canvas später (PX) |

Runtime: Portal-Home = Schul-Pins → persönliche Favoriten → Produkt-Fallback. Persönliche Canvas-Layouts nicht in den Schul-Backup-Monolithen mischen.

---

## 6. Phase P1 – Kalender und Termine (+ Graph)

### Kurz

Zentrale **Termine**-Seite: nächste 14 Tage, typisierte Anlage, Abwesenheit, Sync in Outlook/Gruppenkalender.

### Personas

| Rolle | Kann |
|-------|------|
| Lehrkraft | Interne Termine sehen/anlegen (wenn erlaubt), eigene Abwesenheit starten |
| Direktion / Admin | Öffentliche Termine, ARGE-Tage, Beratungs-Termine, Kalender-Ziele konfigurieren |
| IT | Graph-Scopes, Zielkalender (Shared Mailbox / Gruppe), PA-Fallback |

### MVP-Funktionen

- Liste „Termine in den nächsten 14 Tagen“ aus SharePoint `Schultermine` (oder erweiterte Liste)
- Sichtbarkeit: `öffentlich` / `intern`
- Aktionen:
  - **Internen Termin anlegen** (Titel, ganztägig, Von/Bis, Ort, Beschreibung)
  - **Abwesenheit melden** → Deep-Link / eingebetteter Einstieg in `lehrer-freistellung-planer` (keine Doppel-Logik)
  - **Vollständiger Terminkalender** (Monats-/Jahresansicht)
- Optional im MVP+ (schulkonfigurierbar):
  - ARGE-Tag (Fächer + betroffene Lehrpersonen → Mail an Admin)
  - Bildungsberatungs-Termin (Einrichtung, Link, Mail an Zieljahrgang)

### Graph / Outlook

| Schritt | Verhalten |
|---------|-----------|
| Lesen | Graph `calendarView` auf Schulkalender-, Lehrerkalender-UPN und `/me` |
| Schreiben | Graph `POST/DELETE …/events` in den gewählten Zielkalender |
| Quellen | UPNs aus Tool **Kalender-Postfächer** (Settings im Hub) |
| Rechte | Shared Mailbox braucht Bearbeiten-Freigabe + `Calendars.ReadWrite.Shared` |
| Fallback | PA-Rezept `pa-termine-sync` nur noch optional für Automationen außerhalb des Hubs |

### Nicht im Scope (P1)

Volles Ressourcen-Booking, Vertretungsplan-Ersatz, Schüler-Absenzenbuch.

### Daten / Setup

- Bestehende Liste nutzen oder Spalten nachziehen: `Sichtbarkeit`, `TerminTyp`, `Zielgruppe`, `Ort`, `GraphEventId`
- Tenant-Settings: Zielkalender-UPN, wer darf öffentliche Termine anlegen

### Lizenz

Paket `intranet` und/oder `schulapps`; Feature-ID `termine-hub` (Anzeige: Kalender und Termine).

### Docs / Verkaufssatz

*„Termine nicht nur in Outlook tippen – im Schul-Portal anlegen, freigeben und nach Outlook bringen.“*

---

## 7. Phase P2 – Schulaktivitäten härten

### Kurz

Bestehenden `schulaktivitaeten-planer` an die Referenz-UX angleichen – ohne neues Parallel-Tool.

### Lieferobjekte

| Feature | Beschreibung |
|---------|--------------|
| **Hub-Kacheln** | Anmelden / Meine / Limits / Klassen-Auslastung (wie Screenshot-Übersicht) |
| **Kostenlimits** | Konfigurierbare Tabelle (Veranstaltungstyp → grün/gelb/rot); nur Orientierung, kein Payment |
| **Klassen-Auslastung** | Monatsraster je Klasse: genehmigte/beantragte Aktivitäten als Blöcke |
| **Lehrer-Vorschlag** | Bei Klassenwahl betroffene Lehrkräfte aus Stammdaten/Stundenplan-Hinweisen vorschlagen |
| **Anmerkungen splitten** | Lehrer-Hinweis vs. Admin-Hinweis (falls noch nicht klar getrennt) |
| **Aktionstage** | Optional später: Unterprogramm / eigene View – nicht MVP von P2 |

### Akzeptanz

Lehrkraft versteht Kostenampel und sieht Konflikttage pro Klasse, bevor sie einreicht.

### Lizenz

Bleibt `schulapps` / bestehendes Feature.

---

## 8. Phase P3 – Aufgaben

### Kurz

Interne **Aufgaben** an Personen oder Rollen: „Bitte XY erledigen“ – melden, zuweisen, abhaken, kommentieren.

**Docs:** [`aufgaben.md`](./aufgaben.md) · Familie `betrieb` · Shared Quality Bar (Site/Stammdaten/Calm/Demo) umgesetzt (2026-11).

### Nutzen

Ersetzt lose Teams-Chats / Zettel an Direktion/IT/KV. Aus den Screenshots: eigener Menüpunkt „Aufgaben“ mit klarer Inbox.

### Personas / App-Rollen

| Rolle | Kann |
|-------|------|
| Lehrkraft (Melder) | Aufgabe anlegen, Scope mir / von_mir, Status auf eigenen |
| Direktion / Betrieb-Admin | Schulweite Übersicht (`showAll`), Clear/Demo, System |
| IT | Listen-Assistent / Site (nicht in Lehrer-Primär-Nav) |

### MVP

- Liste `Aufgaben`: Titel, Beschreibung, Zuweisung (Person/Mail), Rolle optional, Priorität, Frist, Status (`offen`/`in_arbeit`/`erledigt`/`abgelehnt`)
- Filter: Mir zugewiesen / Von mir erstellt / Offen (Admin)
- Kommentar-Thread (einfach: Multiline); Actor bei Status/Kommentar
- Demo local-first (`?demo=1` / System-Button); Role-Chips nur `?demoRole=1`
- SharePoint Roundtrip (Pull/Push, soft 403) · Lehrer-Picker (Stammdaten + Entra)
- Benachrichtigung: mailto bei Zuweisung/Erledigt · PA-Rezept `aufgaben-mail`
- Vorlagen-UI: Unterschrift, Unterlagen, Raum prüfen, Elternkontakt

### Abgrenzung

| Nicht Aufgaben | Stattdessen |
|----------------|-------------|
| Technik kaputt (Gerät/Raum) | P4 Technik |
| Genehmigung Exkursion | Schulaktivitäten |
| Lehrer krank | Lehrer-Freistellung |

### Lizenz

Feature `aufgaben` – Paket `betrieb` / `schulapps`.

---

## 9. Phase P4 – Technik-Meldungen

### Kurz

Lehrkräfte melden **defekte / schlechte / austauschwürdige** Technik; IT sieht Inbox mit Priorität und Status.

### Nutzen

Screenshot-Modul „Technik“: Raum, Geräte-Kategorie, Priorität, Offen/Erledigt, Zugangsdaten-Hinweis.

### Personas

| Rolle | Kann |
|-------|------|
| Lehrkraft | Neue Fehlermeldung, eigene offenen sehen |
| IT | Alle Tickets, Status setzen, erledigte Archiv-Ansicht |
| Optional Hausmeister | eigene Queue (Filter Kategorie Gebäude) |

### MVP

**Liste `TechnikMeldungen`:**

| Feld | Beispiel |
|------|----------|
| Titel / Beschreibung | „Beamer flackert“ |
| RaumCode / RaumName | INF3, Physiksaal |
| Kategorie | PC, Beamer, Audio, Netzwerk, Sonstiges |
| Priorität | dringend / hoch / mittel / niedrig |
| Status | offen / in_arbeit / erledigt / abgelehnt |
| MelderEmail / MelderName | aus MSAL |
| Kommentar | optional Thread |
| ErledigtAm | Datum |

UI:

- Ticketliste (Farb-Pills Priorität/Status)
- „Neue Fehlermeldung“
- „Erledigte anzeigen“
- Kachel „Zugangsdaten“ → nur Link/Seite (WLAN-Hinweis), **keine Passwort-Datenbank** im Tool

### Abgrenzung zu Ausgaben / Intune

- **Ausgaben** = wer hat welchen Gegenstand  
- **Technik** = was ist kaputt / melden  
- **Intune-Compliance** = später read-only IT-Check (siehe analyse-04), nicht Lehrkraft-Ticket

### Wiederverwendung

Gleiches Ticket-Pattern wie Aufgaben (Status, Priorität, Kommentar, Mail). Optional gemeinsame Basis-Komponenten.

### Lizenz

Feature `technik-meldungen` – Kandidat Paket `ablauf` oder neues **„Betrieb & Technik“** (mit Raumplan später).

---

## 10. Phase P5 – Raumplan

### Kurz

Schule spielt einen **Raumplan** ein (Bild/PDF/SVG oder strukturierte Raumliste). Lehrkräfte/IT können Räume finden, filtern, optional Belegung sehen und mit Technik/Ausgaben verknüpfen.

### Nutzen

Nicht in den Screenshots prominent – aber starker Schul-Alltagswunsch. Ergänzt schwaches `raeume-ressourcen`.

### Personas

| Rolle | Kann |
|-------|------|
| IT / Verwaltung | Plan hochladen, Raumstammdaten pflegen, Import |
| Lehrkraft | Plan ansehen, Raum suchen, Ticket zu Raum öffnen |
| Sekretariat | Belegungshinweise / Bookings-Link |

### MVP (gestuft)

**Stufe A – Plan + Stammliste**

- Dokumentbibliothek oder Asset: Raumplan (PDF/PNG/SVG)
- SharePoint-Liste `Raeume`: Code, Name, Gebäude, Stockwerk, Kapazität, Typ (Klasse/Fachraum/Büro), optional Graph-Place-Id
- Import CSV (Raumliste der Schule)
- UI: Plan-Viewer + Liste/Suche

**Stufe B – Leichte Interaktion**

- Klickbare Hotspots (optional manuell: Raum → Koordinaten auf Plan)
- Deep-Link „Technik melden“ mit vorausgefülltem Raum
- Anzeige Room-Mailbox / Bookings-Link wenn vorhanden

**Stufe C – Belegung (später)**

- Lesen aus Graph Places / Kalender des Raums
- Kein voller Raum-Buchungsersatz für Untis

### Nicht im Scope

CAD-Editor, automatische PDF-Vektorisierung, Live-Navigations-App.

### Lizenz

Feature `raumplan` – Paket **Betrieb & Technik** oder Ausbau `kommunikation`/`intranet`.

---

## 11. Phase PX – Canvas-Dashboard (Personalisierung)

### Kurz

End-User können Kacheln auf einem Raster **anordnen, skalieren, favorisieren** – Layout wandert mit dem Microsoft-Konto, nicht mit eurer Plattform-DB.

### Lieferobjekte

1. Edit-Modus „Dashboard anordnen“
2. Layout-Schema (x/y/w/h + tileId), tenant-scoped
3. Persistenz: **localStorage Cache** + Sync nach **OneDrive App-Folder** (`approot`) via Graph
4. Fallback ohne Graph-Consent: nur localStorage + Hinweis „nur auf diesem Gerät“
5. Optional später: Schul-Vorlage in SharePoint, die User einmalig übernehmen können

### Akzeptanz

- Lehrer A und Lehrer B sehen unterschiedliche Homes im selben Tenant.
- Gleicher User auf zweitem PC: Layout kommt nach Login/Sync wieder (wenn OneDrive-Pfad aktiv).
- Betreiber-Logs/DB enthalten **keine** Layout-JSONs.

### Lizenz

Teil von `core` / Portal – kein Extra-Preis nötig; Differenzierer in der UX.

---

## 12. Phase P6+ – Weitere Screenshot-Module

Nach dem Kern (P0–P5 / PX), priorisiert nach Nachfrage:

| Modul | Kurzbeschreibung | Aufwand | Hinweis |
|-------|------------------|---------|---------|
| **Q-Handbuch** | Durchsuchbares Prozess-/Begriffshandbuch (SharePoint-Seiten oder Liste A–Z) | mittel | Inhalt liefert die Schule |
| **Unterrichtsmaterial** | Materiallisten je Fach/Klassengruppe („Heft, Buch, Badezeug“) | mittel | Vision-Doc existiert teilweise |
| **Klassengeld-Übersicht** | Jahrgang → Posten inkludiert ja/nein; **ohne** Payment | mittel | Nur Orientierung + Hilfe-Links |
| **Öffentlichkeitsarbeit** | Artikel/Fotos/Leistungen einreichen (CMS-Zulieferung) | mittel | SharePoint News / Bibliothek |
| **Newsletter** | Prüfen: eigener Kanal vs. Elternbrief + Intranet-News | niedrig (Klärung) | Doppelungen vermeiden |
| **Förderbedarf / Förderkurse** | Größeres Pädagogik-Paket | hoch | Bewusst später; Nähe zu Noten/Förderung |

---

## 13. Paketierung (Vorschlag)

| Paket-ID | Label | Inhalt |
|----------|-------|--------|
| *(bestehend)* `schulapps` / `freistellung` / `ablauf` / `intranet` | … | Kalender und Termine, Aktivitäten-Ausbau, Lehrer-Sichtbarkeit |
| `kollegium` *(neu, optional)* | Kollegium & Aufgaben | `aufgaben`, ggf. Kollegiumsumfragen umhängen |
| `betrieb` *(neu, optional)* | Betrieb & Technik | `technik-meldungen`, `raumplan`, ggf. Vertiefung Ausgaben |

Alternativ: alles unter erweitertem `schulapps` – dann Preiserhöhung statt neuer SKUs.

Preis-Orientierung (intern): Aufgaben **90–120 €**, Technik **100–140 €**, Raumplan **120–160 €**/Schuljahr – oder Bundle Betrieb **249 €**.

---

## 14. Plattform-Muster (gleich wie Schul-Apps)

Neue Module = Schema + Listen-Setup + Tool-UI + Feature-Catalog + Docs + Tests.

Wiederkehrende Bausteine:

- Ticket: Status, Priorität, Melder, Zuweisung, Kommentar
- Kalender: SharePoint SoR → Graph Event / ICS / PA
- Portal: Audience Lehrer, Sidebar-Cluster, Schnellzugriff-Links
- Import: CSV für Räume / Limits / Material

---

## 15. Abgrenzung zum Markt

| Konkurrent / System | Wir |
|---------------------|-----|
| WebUntis | Plan & Absenzen bleiben dort; wir: Portal-Termine, Freigaben, Technik, Aufgaben im M365-Tenant |
| Outlook allein | Wir: schulspezifische Formulare + Genehmigung + danach Outlook |
| edu.FLOW / Klasso / Schulmanager | Ähnliche Alltagsmodule; wir: Daten in SharePoint der Schule, einheitlich mit IT-Schultools |
| Virtualschool + Intune | Geräte-MDM bleibt optional später; Technik-Tickets sind bewusst „melden“, nicht MDM |

Verkaufssatz (Entwurf):  
*„IT richtet M365 ein – Lehrkräfte erledigen den Alltag: Termine, Veranstaltungen, Aufgaben und Technik-Meldungen – im gleichen Portal.“*

---

## 16. Offene Entscheidungen

1. Neues Paket `betrieb` / `kollegium` vs. Aufpreis `schulapps`?
2. Kalender und Termine: Graph-Write direkt vs. erst PA-Sync ausbauen?
3. Aufgaben und Technik: eine Ticket-Engine mit Typen oder zwei Tools?
4. Raumplan MVP: nur PDF+Liste (Stufe A) oder gleich Hotspots (Stufe B)?
5. Klassengeld: überhaupt anbieten ohne Payment – oder bewusst streichen?
6. Q-Handbuch: SharePoint-Seiten-Suche vs. eigene Listen-App?
7. Canvas-Sync: **OneDrive approot** (User) vs. SharePoint-Liste (Schule) vs. hybrid?
8. Braucht die Schule Admin-Recht „Layout zurücksetzen“, oder reicht User-only?

---

## 17. Empfohlene Reihenfolge der Umsetzung

1. **P0** Lehrer-Whitelist + Portal-Shell (linke Nav, 3 Spalten) + Schnellzugriff; Schulportal zuerst / IT bewusst  
2. **P1** Kalender und Termine MVP (Liste + interner Termin + Abwesenheit-Link + Kalender-Ziel)  
3. **P2** Aktivitäten: Limits + Klassen-Auslastung  
4. **P3** Aufgaben MVP  
5. **P4** Technik-Meldungen MVP  
6. **P5** Raumplan Stufe A (Import + Viewer)  
7. **PX** Canvas-Dashboard (nach Favoriten-Ausbau; Widgets sobald Termine/Aufgaben da)  
8. Danach: Graph-Vertiefung Termine, Raumplan Stufe B, P6+-Module nach Nachfrage  

---

## 18. Nächster konkreter Schritt

**P0–P5 MVP sind umgesetzt** (localStorage/Demo + Lehrer-Portal-Shell + Lizenzpaket `betrieb` / auch in `schulapps`). Schulportal ist Default; IT-Bereich bewusster Einstieg.  
Als Nächstes: Canvas (PX), feinere Multi-Kalender-Rechte. Kalender und Termine liest/schreibt Outlook (Schul-/Lehrerkalender); SharePoint-Listen-Setup für Aufgaben/Technik/Raumplan ist vorhanden.
