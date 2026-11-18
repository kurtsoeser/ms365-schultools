---
name: schulplaner-tool-upgrade
description: >-
  Upgrades MS365 school apps to Schulaktivitäten-level quality while respecting
  tool families (approval planners, Schularbeiten read/write, Projektwochen/Bookings,
  Einsammler/Elternbrief/Krankmeldung, Wahlen, Betrieb, IT-setup). Use when improving
  Freistellung, Schularbeiten, Projektwochen, Einsammler, Elternbrief, Krankmeldung,
  Schulsprecherwahl, or when propagating a UX/setup fix or SPO Live Sync (SharePoint
  as system of record) across sibling tools; also for Tool-Vorlage, Stammdaten,
  Calm UI, Agent-Upgrade, or „nur Browser-lokal / braucht SharePoint Pull Push“.
---

# Schul-Tool-Upgrade (Familienbewusst)

## Wann

Nutzer will ein **Schul-Tool** verbessern oder eine Verbesserung **auf Geschwister-Tools** übertragen. Referenz-UI: Schulaktivitäten – aber **nicht jedes Tool ist ein Genehmigungs-Planer**.

## Pflicht lesen

1. `docs/tool-upgrade-vorlage.md` (Familien, Shared Bar, Übertragung)  
2. `docs/rollen-vertrag.md` (Lizenz ≠ Portal-Persona ≠ App-Rolle)  
3. `docs/<tool>.md` falls vorhanden  
4. `checklist.md` in diesem Ordner  
5. Referenz: `src/tools/schulaktivitaeten-planer/` + passende Shared-Module  

## Schritt 0 – Familie & Rollen (immer zuerst)

Bestimmen und im Plan festhalten:

| Feld | Beispiel |
|------|----------|
| Familie | `approval-planer` · `register-planer` · `projekt-buchung` · `ablauf-kommunikation` · `wahl-portal` · `betrieb` · `it-setup` |
| Portal-Personas | lehrer / schueler / it |
| App-Rollen | z. B. Lehrer schreibt, Schüler liest; KV; Direktion |
| Spezialfluss | Bookings-Export, Public-Link, Wahl-Phasen, Untis-Ops, … |

Quellen: `dashboard-audience-catalog.js`, `license-feature-catalog.js`, Tool-Docs.

**Stopp:** Domäne nicht umbiegen (kein Fake-Genehmigungsfluss, kein Bookings in Einsammler, kein Schüler-Write bei Read-only).

## Arbeitsmodus

### Upgrade eines Tools
Wellen:
- **A** – Site/Stammdaten/Ownership/Lizenz, Architektur-Split wo nötig  
- **B** – Rollen-Nav + Calm UX passend zur Familie (Schüler-Read beachten)  
- **C** – Familie-Spezial + Demo local-first + Tests/Smoke **pro App-Rolle**

### Übertragung einer Verbesserung
1. Shared-fähig? → `src/shared/`  
2. Sonst gleiche Familie zuerst  
3. Dann nahe Familien  
4. Vorlage/Checkliste aktualisieren, wenn neues Dauer-Pattern  

## Shared Quality Bar (immer)

Site via `school-tool-sites` · Stammdaten hydrate + local-first · kein 403-Wipe · Tokens/`planner-apps-theme` · Calm wo Alltag · Rollenvertrag · Ownership beim Schreiben · Feature-Catalog · **SPO Live Sync (SoR)** – siehe `docs/spo-live-sync.md`.

### Modus SPO Live Sync (Übertragung)

Wenn produktive Daten nur im Browser liegen oder Pull/Push fehlt:

1. Pflicht lesen: `docs/spo-live-sync.md` + `docs/<tool>.md`  
2. **Ist-Scan:** schon `*-graph.js` + Live-CRUD? (z. B. Projektwochen/sRDP oft schon SoR) → nur Lücken, nichts doppelbauen  
3. Wellen **S0→S3** aus `spo-live-sync.md` (Site/Schema → Graph/Sync → UI-Meta → Tests/Smoke)  
4. Referenz Betrieb: `technik-meldungen-graph.js` + `technik-meldungen-sync.js`  
5. Referenz Planer: `srdp-anmeldung-graph.js` / `schulaktivitaeten-planer-graph.js`  
6. Shared-first wenn Boilerplate ≥2 Tools; Domäne/Familie-Spezial (Bookings, Public-Link) nicht vermischen  
7. Abnahme: Item in **Schul-SharePoint** sichtbar; 403 wipe’t nicht; Demo-Seeds nicht pushen  
8. Tracker `docs/spo-live-sync-tracker.md` aktualisieren (Status ✅ + Datum + Notiz; Fortschritt)

## Familien-Kurzregeln

| Familie | Muss | Nicht erzwingen |
|---------|------|-----------------|
| approval-planer | Calm, Drawer, Wizard, oft Shared Kalender, Filter-Multi | — |
| register-planer | Write vs Read nach App-Rolle | KV/Untis nur wenn Konzept |
| projekt-buchung | Export/Sync-Schritt klar | Genehmigungs-Sidebar 1:1 |
| ablauf-kommunikation | Absender calm; Empfänger schlank | Volle Planer-Nav |
| wahl-portal | Phasen, Schüler-Primär | Lehrer-Antrag-Wizard |
| betrieb | Status-Pipeline calm | Kalender Pflicht |
| it-setup | Klare Schritte, IT-Look ok | Lehrer-Calm |

## Anti-Patterns

Alles zu Planern machen · Schüler-Write bei Read · Roh-403-Banner · Role-Chips ohne `?demoRole=1` · Verbesserung nur lokal lassen obwohl Shared · Lizenz/Persona/App-Rolle vermischen.

## Output

- Familie + Rollen genannt  
- Wellen A/B/C (oder Propagationsliste)  
- Shared vs. tool-spezifisch  
- Smoke je Persona  
- Offene Domänenpunkte  
