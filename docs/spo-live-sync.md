# SPO Live Sync – System of Record für Schul-Tools

**Stand:** 2026-11-18  
**Skill:** `.cursor/skills/schulplaner-tool-upgrade/` (Modus „SPO Live Sync“)  
**Vorlage:** [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md) §2 / §7  
**Referenz-Implementierungen:**  
- Planer: `schulaktivitaeten-planer-graph.js`, `srdp-anmeldung-graph.js`, `freistellung-planer-graph.js`  
- Projekt: `projektwochen-graph.js` + `projektwochen-sync.js` (Meta; Bookings-Export extra)  
- Betrieb (frisch): `technik-meldungen-graph.js` + `technik-meldungen-sync.js`

---

## 1. Grundsatz

Die **Schule arbeitet in ihrer SharePoint-Welt**.  
Browser-localStorage ist **Cache / Offline / Demo**, nicht die Wahrheit für produktive Vorgänge.

| Schicht | Rolle |
|---------|--------|
| SharePoint-Liste(n) | **System of Record** (SoR) |
| App | Formular + Rollen-UI + Pull/Push |
| localStorage | Cache, Demo (`localDemoOnly` / `seedTag`) |
| IT-Listen-Assistent | Liste + Spalten anlegen, Site-URL merken |

**Stopp:** Domäne nicht umbiegen (kein Fake-Genehmigungsfluss, Bookings nur wo Konzept).

---

## 2. Muss-Pattern (jedes datenhaltende Schul-Tool)

1. **Site** über `resolveSchoolToolSiteUrl('<kind>')` + Tool-Settings (kein Hardcode).  
2. **Schema** (`*-schema.js`) = Spalten der SPO-Liste; IT-Assistent `sharepoint-liste-<tool>`.  
3. **Graph** (`*-graph.js`): resolve Context, map From/To, list/create/update/delete via `spo-graph-shared` (`window.ms365SpoGraph`).  
4. **Sync** (eigenes Modul oder Entry):  
   - Start → **Pull**  
   - Schreiben → lokal sofort → **Push**  
   - `spoItemId` (o. Ä.) merken  
5. **403/Fehler** → ruhiger Hinweis, **kein Wipe** von Demo/Cache.  
6. **Demo-Seeds** nicht blind nach SPO pushen; „Demo laden“ bleibt local-first.  
7. **Ownership/Actor** beim Write setzen.  
8. **Tests** für Mapping + Scope; Smoke mit Site.

---

## 3. Ist-Scan (vor Umbau)

| Frage | Ja → | Nein → |
|-------|------|--------|
| Hat das Tool schon `*-graph.js` + Live-CRUD? | Nur Lücken schließen | Graph + Sync bauen |
| Ist SPO schon SoR laut `docs/<tool>.md`? | Docs + Smoke prüfen | SoR festschreiben + verdrahten |
| Nur Bookings/Export extra? | CRUD und Export trennen | – |
| Public/Magic-Link? | Token-Form schreibt trotzdem SPO (oder Backend) | – |

---

## 4. Wellen (nur SPO-Live)

| Welle | Inhalt |
|-------|--------|
| **S0** | Ist-Scan SoR; Site-Kind; Schema ↔ Liste |
| **S1** | Graph map + CRUD; Sync Pull/Push; 403-Policy |
| **S2** | UI: Sync-Meta System/IT; ruhige Fehler; Demo nicht pushen |
| **S3** | Tests Mapping; Docs; Smoke Melder/Lehrer/Admin + IT-Liste |

---

## 5. Propagationsreihenfolge

1. Gleiche Familie (z. B. Betrieb: Aufgaben → Ausgaben → Raumplan)  
2. Ablauf-Apps ohne Public-Sonderfall  
3. Wahl-Portale (Schüler-Write → SPO)  
4. Planer nur wenn CRUD noch lokal ist (viele sind schon live)  
5. **Nie** Bookings in Nicht-Buchungs-Tools erzwingen  

Shared-first: wiederkehrende Hilfen nach `src/shared/` ziehen, wenn ≥2 Tools denselben Boilerplate brauchen.

---

## 6. Anti-Patterns

- Nur localStorage für produktive Tickets/Anmeldungen/Anträge  
- Pull ohne Push (oder umgekehrt)  
- Demo-Seed nach SPO schreiben und echte Daten vermischen  
- 403 löscht lokalen Stand  
- IT-Listen-Link in Lehrer-Primär-Nav  
- „Wir machen später Sync“ ohne Site/Schema-Fundament  

---

## 7. Agent-Prompts & Tracker

- **Tool-für-Tool-Liste mit Status + kopierbaren Prompts:** [`spo-live-sync-tracker.md`](./spo-live-sync-tracker.md)  
- Vorlage/Skill: [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md) §12 · Skill `schulplaner-tool-upgrade`

Nach jedem Tool den **Tracker** aktualisieren (Status/Datum/Notiz).
