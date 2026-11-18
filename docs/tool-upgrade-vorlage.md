# Vorlage: Schul-Tools auf „Schulaktivitäten-Niveau“ bringen

**Stand:** 2026-11-18  
**Referenz-UI/UX:** Schulaktivitäten-Planer (`tools/schulaktivitaeten-planer.html`, `src/tools/schulaktivitaeten-planer/`)  
**Rollenvertrag:** [`rollen-vertrag.md`](./rollen-vertrag.md)  
**Zweck:** Einheitliche Vorgangsweise für Setup, SharePoint/Stammdaten, Layout/Design, UX – **bei klarer Anpassung an Tool-Familien**. Nicht jedes Tool ist ein Genehmigungs-Planer; trotzdem sollen Verbesserungen (Calm UI, Site-Hydrate, Demo, Filter, …) **mit wenig Aufwand** auf Geschwister-Tools übertragbar sein.

**Cursor:** Skill `.cursor/skills/schulplaner-tool-upgrade/` · Rule `.cursor/rules/schulplaner-tool-upgrade.mdc`

---

## 0. Grundidee

Ein Schul-Tool ist **kein Admin-Dashboard**, sondern ein **Arbeitsplatz** mit klar getrennten Schichten:

| Schicht | Frage | Beispiele |
|--------|--------|-----------|
| **Lizenz** | Hat die *Schule* das Modul? | `license-feature-catalog` / License-API |
| **Portal-Persona** | Wer bin *ich* im Portal? | IT · Lehrkraft · Schüler/in |
| **App-Rolle** | Was darf ich *in diesem Tool*? | Lehrer schreibt / Schüler liest · KV · Direktion · Admin |

Siehe [`rollen-vertrag.md`](./rollen-vertrag.md) – **nicht vermischen**.

Mentale Welten in der App:

| Welt | Ziel |
|------|------|
| Lehrkraft (+ KV als Fähigkeit) | Handeln, wenig Mental Load |
| Schüler/in (wo vorgesehen) | Lesen / eigene Anträge / Wahl – selten Admin |
| Leitung / Admin-Inbox | Entscheiden, auswerten, nacharbeiten |
| System / IT | Site, Listen, Flows, Diagnose, Demo |

**Kurzformel:**  
Familie erkennen → Shared UI/Setup wiederverwenden → App-Rollen & Spezialflüsse tool-spezifisch lassen → Demo local-first → Verbesserungen in Shared ziehen, wo möglich.

---

## 1. Tool-Familien (nicht alles ist „Planer“)

Vor jedem Upgrade: **Familie bestimmen**. UI-Shell und Setup sind oft gleich; Domäne und Rechte nicht.

### 1.1 Familie `approval-planer` (Genehmigung + Kalender)

**Beispiele:** Schulaktivitäten, Freistellung Schüler, Freistellung Lehrkräfte  
**Kern:** Antrag → (KV) → Direktion → Liste/Kalender → optional Ops-Nacharbeit  
**UI-Pflicht:** Calm-Dashboard, Drawer, Guided Create, Shared Kalender, Mehrfachfilter  
**Referenz:** `schulaktivitaeten-planer`

### 1.2 Familie `register-planer` (Eintragen + Lesen nach Rolle)

**Beispiele:** Schularbeiten-Planer  
**Kern:** Lehrkräfte **schreiben** Termine; Schüler oft **nur lesen**; Leitung/Admin Übersicht  
**Besonderheit:** App-Rollen `lehrer` / `schueler` / `admin` – Schreib-UI für Schüler ausblenden, Read-Views calm halten  
**Nicht erzwingen:** Untis-Ops-Banner oder KV-Freigabe, wenn das Fachkonzept anders ist  
**Referenz-Docs:** `docs/schularbeiten-planer.md`

### 1.3 Familie `projekt-buchung` (Planung + externes Booking)

**Beispiele:** Projektwochen, Elternsprechtag-Bookings  
**Kern:** Planung/Zuordnung in der App; **Export / Abbildung** nach Microsoft Bookings (o. Ä.)  
**UI-Pflicht:** Calm Setup, klare „Export/Sync“-Aktion in System oder Abschluss-Schritt – kein Pseudo-Genehmigungsfluss erzwingen  
**Referenz-Docs:** `docs/projektwochen.md`

### 1.4 Familie `ablauf-kommunikation` (Einsammeln / Informieren)

**Beispiele:** Einsammler, Elternbrief, Krankmeldung, Kollegium-Umfragen  
**Kern:** Absender (Lehrkraft/Leitung) startet Aktion; Empfänger antworten oder bestätigen (Link/Form)  
**UI-Pflicht:** Ruhige Absender-Inbox; Empfänger-Oberflächen schlank (eigene HTML-Seiten ok)  
**Nicht erzwingen:** Sidebar wie Planer, wenn das Tool ein kurzer Wizard + Liste ist – Calm Tokens und Site/Stammdaten trotzdem  
**Referenz-Docs:** `docs/einsammler.md`, `docs/elternbrief.md`, `docs/krankmeldung.md`

### 1.5 Familie `wahl-portal` (Mitbestimmung / Anmeldung)

**Beispiele:** Schulsprecherwahl, Wahlfach-Zuteilung, SRDP-Anmeldung  
**Kern:** Zeitraum, Berechtigte, Abgabe; oft Schüler als Primärnutzer  
**UI-Pflicht:** Klare Phasen (offen/geschlossen), Trust/Privacy, Admin-Auswertung getrennt  
**Referenz-Docs:** `docs/schulsprecherwahl.md`, `docs/wahlfach-zuteilung.md`, `docs/srdp-anmeldung.md`

### 1.6 Familie `betrieb` (Tickets / Inventar / Räume)

**Beispiele:** Aufgaben, Technik-Meldungen, Raumplan, Ausgaben  
**Kern:** Melden → Bearbeiten → Erledigt; Inventar/Ressourcen  
**UI-Pflicht:** Status-Board calm; IT-Setup getrennt  

### 1.7 Familie `it-setup` (Listen / Playbooks / PA)

**Beispiele:** `sharepoint-liste-*`, Playbooks, Power-Automate-Rezepte  
**Kern:** Nur Schul-IT; darf „technischer“ aussehen (dashed IT-Panels ok)  
**Upgrade-Fokus:** klare Schritte, Site-URL, Erfolgsfeedback – nicht Lehrer-Calm erzwingen  

### Familien-Matrix (Orientierung)

| Tool (Auszug) | Familie | Typische App-Rollen | Spezial |
|---------------|---------|---------------------|---------|
| Schulaktivitäten | approval-planer | Lehrer, KV*, Admin | Untis-Nacharbeit, ICS/Outlook |
| Freistellung Schüler | approval-planer | Schüler, KV, Direktion | Schüler beantragt |
| Freistellung Lehrer | approval-planer | Lehrer, Direktion | anderes Schema |
| Schularbeiten | register-planer | Lehrer schreibt, Schüler liest, Admin | Konflikt/Überblick |
| Projektwochen | projekt-buchung | Lehrer, Admin | Bookings-Export |
| Einsammler / Elternbrief / Krankmeldung | ablauf-kommunikation | Absender + Empfänger-Link | Public/Token-Formulare |
| Schulsprecherwahl / Wahlfach / SRDP | wahl-portal | Schüler (+ Admin) | Phasen, Quoren |
| Aufgaben / Technik / Raumplan / Ausgaben | betrieb | Melder / Bearbeiter | Status-Pipeline |
| Schüler-Stammdaten-Portal | hybrid (ablauf + betrieb + it-setup) | Sekretariat/IT + Eltern-PIN | Gateway/SharePoint, Status→Sokrates; kein Genehmigungs-Planer |
| Listen-Assistenten / Playbooks | it-setup | IT | Setup only |

\* KV = Capability aus Stammdaten/`headEmail`, oft keine eigene Portal-Persona.

Portal-Sichtbarkeit: `src/shared/dashboard-audience-catalog.js` (`DASHBOARD_PORTAL_SECTIONS_BY_VIEW`).  
Lizenz-Module: `src/shared/license-feature-catalog.js`.

---

## 2. Was ist überall gleich (Shared Quality Bar)

Diese Punkte gelten **familienübergreifend** – hier lohnt Übertragung von Verbesserungen:

1. **Site-URL** über `school-tool-sites` + Lizenz, nicht hardcoden  
2. **Stammdaten** live hydratisieren + Meta + local-first Fallback  
3. **403/SPO-Fehler** wipe’t Demo/Cache nicht; kein Roh-Graph-Banner für Endnutzer  
4. **Design-Tokens** `app.css` + `planner-apps-theme.css`; Tool-Prefix-CSS  
5. **Calm UX** wo Alltag: Action-Banner / Quiet Empty; dashed nur IT  
6. **Rollenvertrag** einhalten; Demo-Chips nur `?demoRole=1`  
7. **Ownership / Actor** beim Schreiben setzen (wer hat angelegt/geändert)  
8. **Fach-Labels** in UI, Codes SPO-intern  
9. **Lizenz-Gate** + Feature-Catalog-Eintrag  
10. **Tests** für Logic/Rollen; Smoke Lehrer / Schüler / Admin je nach Audience  
11. **SPO Live Sync (SoR):** produktive Daten **Pull + Push** in die Schul-SharePoint-Liste; localStorage nur Cache/Demo – Details [`spo-live-sync.md`](./spo-live-sync.md) · **Tracker:** [`spo-live-sync-tracker.md`](./spo-live-sync-tracker.md)  

---

## 3. Was bewusst tool-spezifisch bleibt

| Thema | Nicht blind von Schulaktivitäten kopieren |
|-------|-------------------------------------------|
| Statusfluss | Nur wenn Genehmigung fachlich existiert |
| Shared Kalender | Nur wenn Termine zentral sind |
| KV-Nav „Meine Klasse“ | Nur wenn KV-Arbeit im Konzept steckt |
| Schüler-Write | Nur wenn Persona/App-Rolle Schreiben erlaubt |
| Bookings/Export | Nur Familie `projekt-buchung` (+ ähnliche) |
| Public Magic-Link | Ablauf-Familie; getrenntes HTML ok |
| Untis/Ops-Flag | Nur wenn echte Nacharbeit existiert |
| Wizard mit 4 Schritten | Nur bei komplexen Anträgen; Einsammler darf kürzer sein |

**Regel:** Shared Shell + Setup + Calm Patterns **ziehen**; Domänenfluss aus `docs/<tool>.md` **respektieren**.

---

## 4. Architektur-Split (wo das Tool wächst)

Für größere Alltags-Apps (Planer, Register, Betrieb):

| Schicht | Datei | DOM? |
|---------|-------|------|
| Mount | `tools/<tool>.html` | ja |
| Entry | `*-planer.js` / `*.js` | ja |
| State | `*-state.js` | nein* |
| Logic | `*-logic.js` | **nein** |
| Schema | `*-schema.js` | nein |
| Graph | `*-graph.js` | nein |
| UI / Detail / Wizard | `*-ui.js` … | Render |
| Demo | `*-demo-data.js` | nein |
| CSS | `*-….css` mit Prefix | – |

Kleine Ablauf-Tools dürfen schlanker bleiben – trotzdem Logic von UI trennen, sobald Validierung/Rollen wachsen.

---

## 5. Setup: SharePoint & Stammdaten

1. `resolveSchoolToolSiteUrl('<toolKey>')`  
2. Lizenz-`ToolSiteUrl` / Bootstrap  
3. Listen laut Schema; Codes statt Lookups  
4. `hydrateStammdatenFromLiveSite` + `stammdatenLive`  
5. System-View nur Staff/IT  
6. Optional: Listen-Assistent `sharepoint-liste-<tool>.html` und PA-Setup als **Geschwister**, nicht in die Lehrer-Nav

---

## 6. Layout & UX (an Familie anpassen)

### Shell
`tm-page` + Prefix-Page; Sidebar nur wenn Nav ≥ ~4 Arbeits-Views.  
Kurze Formulare (Krankmeldung öffentlich) brauchen oft **kein** Planer-Sidebar – trotzdem gleiche Tokens.

### Rollen-UI
| Persona/Rolle | Erwartung |
|---------------|-----------|
| Lehrer (Write) | Primär-CTA, Filter, eigene Einträge |
| Schüler (Read) | Keine „Neue …“-Buttons, klare Read-only-Liste/Kalender |
| Schüler (Write, z. B. Freistellung) | Guided Antrag, ruhige Statusanzeige |
| KV | Zusätzliche Arbeit („Meine Klasse“ / Stellungnahme) |
| Direktion/Admin | Entscheidungen / Alle / Auswertung |
| IT | Systemadministration |

### Calm Patterns
Action-Banner · Ops-Banner · Quiet Empty · Drawer statt Modal-Spam · Mehrfachfilter wo Listen lang sind · Shared Kalender wenn Termine.

---

## 7. Verbesserungen übertragen (Wartungs-Traum)

Wenn du **eine** Verbesserung hast (z. B. bessere Nacharbeit-Box, Filter-Multi, 403-Demo-Schutz):

### Schritt A – Shared vs. Tool
- Liegt die Lösung in `src/shared/`? → **einmal** fixen, alle profitieren.  
- Liegt sie nur im Tool-CSS/UI? → Pattern in diese Vorlage + Skill-Checkliste aufnehmen, dann **Geschwister der gleichen Familie** nachziehen.

### Schritt B – Propagationsreihenfolge
1. Gleiche Familie (z. B. beide Freistellungen)  
2. Nahe Planer (`approval-planer` ↔ `register-planer`)  
3. Andere Alltags-Apps mit Sidebar  
4. Ablauf/Wahl nur wenn das Pattern passt (nicht Bookings in Einsammler zwingen)

### Schritt C – Agent-Prompt
```text
Übertrage die Verbesserung „<Name>“ aus Schulaktivitäten
auf die Familie approval-planer (Freistellung Schüler + Lehrer).
Shared zuerst prüfen. Keine Domänenänderung. Skill schulplaner-tool-upgrade.
```

**SPO Live Sync (häufigster SoR-Fix):**
```text
Propagiere „SPO Live Sync (SoR)“ von Technik-Meldungen auf <Ziel-Tool oder Familie>.
Skill schulplaner-tool-upgrade · Docs docs/spo-live-sync.md + docs/<tool>.md
Ist-Scan zuerst (schon live?). Site + Schema + Graph Pull/Push. 403 kein Wipe.
Domäne nicht umbiegen. Am Ende Smoke + IT-Listen-Check.
```

### Schritt D – Kurz-Abnahme
Pro Ziel-Tool: Smoke der betroffenen Personas (nicht nur Admin) **und** ein Live-Item in SPO sichtbar.

---

## 8. Demo local-first

1. Seed-Paket + Seed-Tag  
2. `localDemoOnly`  
3. Persona passend zur Demo-Story (Lehrer/Schüler)  
4. Kalender/Listen auf sinnvolle „heute“-Daten  
5. SPO optional; Fail → lokal bleiben, ruhiger Toast  
6. Reset nur Seed-Zeilen  

---

## 9. Shared Modules (reuse)

| Modul | Zweck |
|-------|--------|
| `ms365-calendar-shell.js/.css` | Agenda/Woche/Monat |
| `school-tool-sites.js` | Tool-Site-URL |
| `tenant-school-site-bootstrap.js` | Lizenz → Site |
| `school-stammdaten-spo-lists.js` | Live Klassen/Lehrer |
| `planner-apps-theme.css` | Tokens |
| `frontend-planner-chrome-policy.js` | IT vs Lehrer Chrome |
| `app-role-switcher-policy.js` | Wann Role-Switcher |
| `dashboard-audience-catalog.js` | Portal-Sichtbarkeit |
| `license-gate*.js` + Feature-Catalog | Module |
| `msal-auth-ui`, `pin-gate`, `app-data-v2`, `spo-graph-shared` | Basis |

---

## 10. Golden Path (Agent)

1. **Familie + App-Rollen** aus Docs/Katalog bestimmen  
2. Shared Quality Bar (§2) gegen Ist-Zustand scannen  
3. Wellen: **A** Setup/Stammdaten/Ownership → **B** Rollen-Nav/Calm UX → **C** Familie-Spezial (Kalender / Bookings / Public-Form / Read-only Schüler) + Demo  
4. Kein Feature-Creep: Domäne aus `docs/<tool>.md`  
5. Tests + Smoke je Persona  
6. Wenn etwas Allgemeines rauskommt → Shared oder Vorlage aktualisieren (§7)

Detail-Checkliste: `.cursor/skills/schulplaner-tool-upgrade/checklist.md`

---

## 11. Anti-Patterns

- Alle Tools zu Genehmigungs-Planern machen  
- Schüler-Write-UI für reine Read-Rollen  
- Bookings-Export in Tools ohne Buchungsbedarf  
- Planer-Sidebar in einzeilige Public-Formulare zwingen  
- Gestrichelte Lehrer-Panels / Roh-403-Banner  
- Role-Chips produktiv ohne `?demoRole=1`  
- Verbesserung nur in einem Tool lassen, obwohl sie Shared gehört  
- Lizenz / Portal-Persona / App-Rolle vermischen  

---

## 12. Agent einsetzen

```text
Upgrade <Tool> auf Qualitätsniveau Schulaktivitäten.
Familie: <approval-planer|register-planer|projekt-buchung|ablauf-kommunikation|wahl-portal|betrieb|it-setup>
Skill: schulplaner-tool-upgrade
Docs: docs/tool-upgrade-vorlage.md + docs/rollen-vertrag.md + docs/<tool>.md
Shared Quality Bar inkl. SPO Live Sync (docs/spo-live-sync.md).
Shared zuerst. Domäne nicht umbiegen. Am Ende Smoke für die App-Rollen des Tools.
```

**Wartung einer Verbesserung über mehrere Tools:**

```text
Propagiere „<Verbesserung>“ von <Quell-Tool> auf <Ziel-Liste oder Familie>.
Skill schulplaner-tool-upgrade § Übertragung. Shared-first.
```

**Nur SharePoint-Live (ohne UX-Komplett-Upgrade):**

```text
SPO Live Sync für <Tool>.
Referenz: src/tools/technik-meldungen/ (graph+sync) bzw. passende Planer-Graph.
Docs: docs/spo-live-sync.md + docs/<tool>.md
Skill: schulplaner-tool-upgrade · Wellen S0→S3
Ist-Scan: schon SoR? Lücken nur schließen. Domäne nicht umbiegen.
```
