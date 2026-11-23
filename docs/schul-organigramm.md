# Schul-Organigramm (Phase 0 – Spec)

**Ziel:** Zweiter Pflegeweg für die Stammdaten – die **gesamte Schule** als Organigramm (nicht nur Verwaltung). Visuell strukturieren, bearbeiten, speichern; Listen-Tabs bleiben parallel gültig.

**Status:** Phase 0–6 abgeschlossen. Weiterentwicklung: **Schul-Landkarte** (verschachtelbare Areas, Canvas) – siehe [schul-organigramm-landkarte.md](schul-organigramm-landkarte.md) (Schema `version: 2`).

---

## 1. Abgrenzung zu bestehenden Werkzeugen

| Werkzeug | Aufgabe | Stammdaten schreiben? |
|----------|---------|------------------------|
| **Schul-Organigramm** (dieses) | Schule **organisieren & pflegen** (Rollen, ARGEs, Klassen, …) | **Ja** (Listen + Layout) |
| Stammdaten-Tabs (`tenant.html`) | Listen-/Tabellen-Pflege | Ja |
| Verwaltung → Organigramm-Tab | Embed des Schul-Organigramms; Legacy-Layout archiviert | Ja (über Schul-Organigramm) |
| [Datenlandkarte](datenlandkarte.md) | Listen/Quellen als Blöcke + Sync-Linien | Nein (nur Layout lokal) |
| [Schuldaten-Karte](schulgraph.md) | Beziehungsgraph erkunden | Nein |

Kurz: Datenlandkarte = „welche Listen gibt es?“, Schulgraph = „wer hängt wie zusammen?“, **Schul-Organigramm = „so ist die Schule aufgebaut – und hier pflege ich sie“**.

---

## 2. Einstiege (festgelegt)

1. **App-Header:** Button „Organigramm“ links neben „Stammdaten“ (`#ms365HeaderSchulOrganigramm`)
2. **Dashboard-Karte:** `data-tool-id="schul-organigramm"` → `tools/schul-organigramm.html`
3. **Stammdaten-Tab:** „Organigramm“ → `#organigramm` / `tabMainOrganigramm` (Embed)

Alle drei zeigen dieselbe Logik (`mountSchulOrganigramm`).

---

## 3. Eine Quelle der Wahrheit

Kanonisch: `localStorage` **`ms365-schooltool-data-v2`**.

| Feld | Inhalt |
|------|--------|
| `core.administration` / `adminRoles` / `admin` | Verwaltungsrollen + Personen |
| `core.arges` | ARGEs / Fachschaften (+ Leitung) |
| `core.teachers` | Lehrkräfte |
| `years.*.classes` | Klassen (+ KV) |
| `years.*.students` / `studentCouncil` | Schüler / Vertretung |
| `core.sga` | SGA |
| `core.verwaltungAudienceGroups` / `adminAudienceMemberships` | Zielgruppen |
| **`core.schulOrganigramm`** | **Layout + Nesting + M365-Bindung** der Schul-Landkarte |
| `core.verwaltungOrganigramm` | Legacy-Archiv (nach Migration nicht mehr aktiv bearbeitet) |

Spiegel: `ms365-tenant-settings-v1` (Komfort).  
Browser-Backup & IT-Sicherungsbibliothek speichern den **vollen** v2-Container (kein Extra-Key nötig).

**Prinzip:** Knoten **referenzieren** Listen (`refKind` + `refKey`). Personenlisten liegen **nicht** doppelt im Layout.  
**Zielbild (Kanon):** [schul-struktur-kanon.md](schul-struktur-kanon.md) – Landkarte, Listen, Backup und Importe (WebUntis / Bildungsportal) teilen **eine** Wahrheit.

---

## 4. Layout-Schema

```text
schulOrganigramm = {
  version: 1,
  nodes: OrgNode[]
}

OrgNode = {
  id: string,                 // stabil, z. B. sec_schulleitung, role_SEK
  type: NodeType,
  label: string,              // Anzeige (kann von Stammdaten-Namen abweichen, bis Sync)
  parentId: string | null,
  order: number,
  tone?: string,              // lead | staff | arge | class | council | …
  refKind: RefKind,
  refKey: string              // z. B. Rollencode SEK, ARGE-Kürzel MATHE, Klassen-Code 1A
}
```

### 4.1 NodeType

| `type` | Bedeutung | typischer `refKind` |
|--------|-----------|---------------------|
| `section` | Bereich / Überschrift (Schulleitung, ARGEs, …) | `none` |
| `role` | Verwaltungsrolle | `adminRole` |
| `arge` | ARGE / Fachschaft | `arge` |
| `class` | Klasse | `class` |
| `council` | Schülervertretung (Person) | `studentCouncil` |
| `sga` | SGA-Eintrag | `sga` |
| `custom` | freier Knoten (nur Organigramm) | `none` oder `audience` |

### 4.2 RefKind → Stammdaten

| `refKind` | Lookup |
|-----------|--------|
| `adminRole` | `adminRoles[].code` (+ Personen über `admin[].role`) |
| `arge` | `arges[].code` |
| `class` | `classes[].code` (aktuelles Schuljahr) |
| `studentCouncil` | E-Mail oder stabiler Key in `studentCouncil` |
| `sga` | E-Mail / Key in `sga` |
| `audience` | `verwaltungAudienceGroups[].id` |
| `none` | kein Listen-Bezug (Section/Custom) |

### 4.3 Seed-Bereiche (Standard)

| Section-`id` | Label | Kinder aus |
|--------------|-------|------------|
| `sec_schulleitung` | Schulleitung | Rollen mit Tier `schulleitung` · `refKind: audience` / `schulleitung` |
| `sec_verwaltung` | Verwaltung (Personal) | übrige Admin-Rollen · `refKind: audience` / `verwaltung` |
| `sec_aud_<va-*>` | Custom-Verwaltungsgruppen | aus `verwaltungAudienceGroups` (z. B. Erweiterte Schulleitung, Schulwarte) |
| `sec_arge` | ARGEs / Fachschaften | `arges` |
| `sec_lehrer` | Lehrkräfte | Sammelgruppe Lehrer |
| `sec_schueler` | Schüler | Sammelgruppe Schüler · darunter Jahrgänge |
| `sec_jahrgang_<n>` | *n*. Jahrgang | Klassen mit führender Stufe *n* im Kürzel (`refKind: jahrgang`) – Basis für spätere M365-Verteiler |
| `sec_vertretung` | Schülervertretung & SGA | `sga_board` mit **Schülervertretung** (3 Funktionen: Schulsprecher, STV1, STV2) · Lehrer · Eltern |

Hierarchie Schüler: **Schüler → Jahrgang → Klasse**. Legacy-`sec_klassen` wird beim Sync in Jahrgänge überführt.

Custom-Bereiche werden im Organigramm („+ Bereich“) **und** unter Stammdaten → Verwaltung (`verwaltungAudienceGroups`) geführt.

---

## 5. Lesen / Schreiben / Sync

```text
Öffnen
  → load Stammdaten
  → wenn schulOrganigramm leer und verwaltungOrganigramm vorhanden: Migration + Sync
  → wenn schulOrganigramm leer: seed aus Listen
  → sonst: sync (fehlende Refs ergänzen, bestehende Knoten behalten)

Bearbeiten (ab Phase 2–4)
  → Pop-up ändert Stammdaten-Liste (admin / adminRoles / arges / classes / studentCouncil / sga)
  → optional Label/parentId im Layout
  → ms365TenantSettingsSave (normalize + setCoreFromTenantSettings)

Teil-Saves (Autosave Stammdaten-UI)
  → schulOrganigramm / verwaltungOrganigramm nie stillschweigend leeren
    (Preserve-Semantik in tenant-settings-core)
```

**„Aus Stammdaten neu aufbauen“:** ersetzt das Layout durch frischen Seed (bewusst destruktiv für die Anordnung).

---

## 6. UI-Konzept (Phasen)

| Phase | Inhalt |
|-------|--------|
| **0** | Diese Spec |
| **1** | Einstiege + Lesen/Seed/Sync + Persistenz + Filter/Farben (Anzeige) – **erledigt** |
| **2** | Schreiben: Rollen/Personen (Pop-up); Farben = Besetzung – **erledigt** |
| **3** | ARGEs (Leitung) + Lehrer-Bezüge – **erledigt** |
| **4** | Klassen (KV), Schülervertretung/SGA – **erledigt** |
| **5** | Migration `verwaltungOrganigramm` → `schulOrganigramm`, Verwaltung-Builder vereinfachen – **erledigt** |
| **6** | Polish: Filter, Druck, DnD, Abgleich-Zähler IT-Sicherung – **erledigt** (erster Abschluss) |

Interaktion (Ziel ab Phase 2): Doppelklick / Stift → Detail-Pop-up; DnD für Bereiche; immer mehrere Personen pro Rolle möglich (Farbe = tatsächliche Besetzung).

---

## 7. Dateien

| Pfad | Rolle |
|------|--------|
| `docs/schul-organigramm.md` | Spec (dieses Dokument) |
| `tools/schul-organigramm.html` | Vollbild-Tool |
| `src/tools/schul-organigramm/schul-organigramm-logic.js` | Seed, Normalize, Sync, Migration Verwaltung → Schul |
| `src/tools/schul-organigramm/schul-organigramm-ui.js` | Mount / Render |
| `src/tools/schul-organigramm/schul-organigramm-edit.js` | Entry/Re-Export + `createSchulOrganigrammEditor` (dünn; öffentliche API) |
| `src/tools/schul-organigramm/schul-organigramm-edit-*.js` | Edit-Split: Utils/HTML/People/Ops/Find/Overlay + UI-Factories (Rollen, ARGE, Klasse, Personen, Area/M365, Canvas) |
| `src/tools/schul-organigramm/schul-organigramm.css` | Styles |
| `src/shared/tenant-schul-organigramm-embed.js` | Tab in `tenant.html` |
| `tests/schul-organigramm-logic.test.mjs` | Seed/Sync-Tests |

Verwandt (Legacy-Logik bleibt für Tests/Archiv):  
`src/tools/verwaltung/verwaltung-organigramm-logic.js`

---

## 8. Abgrenzung / Zielbild

- **Kein zweites Personenmodell** neben `admin` / Jahrgangs-Buckets – Importe schreiben in die Listen, Sync aktualisiert die Landkarte.
- **M365 pro Area:** Match/Create (Gruppe, SG, MESG, Verteiler) + Postfach-Match am Knoten (`m365`) – siehe Landkarte-Spec Phase E.
- Standardgruppen Lehrer/Schüler/ARGE/Klasse bleiben der Einrichtungspfad; freie Areas sind flexibel.
- Live-Graph-Mitgliedschaften im Zoom (Fokus) sind Anzeige; kanonische Pflege bleibt Listen + Area-Editor.
- Roadmap BIP/WebUntis-Zuordnung: [schul-struktur-kanon.md](schul-struktur-kanon.md).

### 8.1 Sync-Semantik (asymmetrisch, bewusst)

Kein klassischer Zwei-Wege-Realtime-Sync. Stattdessen:

| Richtung | Was passiert |
|----------|----------------|
| **Listen → Landkarte** | Materializer (`syncSchulOrganigrammFromAppData` / `writeSchoolData`) ergänzt Knoten, Labels, Eltern-/Chat-Bäume; Canvas-Orte bleiben |
| **Organigramm → Listen** | Nur über **explizite Editor-Commands** (Klasse, KV, ARGE, Rolle, …) |
| **Layout speichern** | `writeSchulOrganigramm` schreibt `core.schulOrganigramm`; Listen bleiben unverändert |

**Dual-Pflege:** Stammdaten-Tabs und Organigramm sind zwei Sichten derselben Entitäten. Tabs besitzen Listenfelder; Organigramm besitzt Hierarchie, Nesting, Canvas und M365-Bindings. Event `ms365-school-data-written` + `ms365-school-year-changed` refreshen die Organigramm-UI.

**Löschen / Umbenennen (P3 – destruktive Policies):**

| Aktion | Verhalten |
|--------|-----------|
| Klasse/ARGE im Editor oder als Entity-Kachel löschen | **Listen + Landkarte** (inkl. Eltern-/Chat-Spiegel, Kursteams unter der Klasse). Schüler bleiben erhalten und behalten die alte Kennung (Hinweis im Dialog). |
| Eltern-/KlassenChat-/Kursteam-Kachel „löschen“ | **Blockiert** – abgeleitete Spiegel, Sync legt sie neu an. |
| Custom-Area / optionaler Builtin von der Landkarte | Nur Layout (`removedAreaIds`); keine Listen-Entität. |
| Klasse in Tabs löschen | Materializer entfernt Phantom-Knoten aus der Landkarte. |
| Klassenkürzel umbenennen (Editor) | Propagiert auf `students.klasse`, `studentCouncil`, `classChats`, exakte Belegungszeilen und alle Klassen-Refs/IDs in der Landkarte. |

---

## 9. Phase-1-Abnahme

- [x] Header-Button „Organigramm“ neben Stammdaten
- [x] Dashboard-Karte `schul-organigramm`
- [x] Stammdaten-Tab `#organigramm` (Embed)
- [x] Seed aus Listen (Rollen, ARGEs, Klassen, Vertretung/SGA)
- [x] Sync: fehlende Knoten ergänzen, Labels aus Stammdaten aktualisieren
- [x] Persistenz `core.schulOrganigramm` (Preserve bei Teil-Saves)
- [x] Bereichsfilter, Legende, Besetzungsfarben (nur Anzeige)
- [x] „Neu aufbauen“ mit Bestätigung
- [x] IT-Sicherungs-Abgleich zählt Schul-Organigramm-Knoten

---

## 10. Phase-2-Abnahme

- [x] Doppelklick / Stift auf Rollen-Karte → Pop-up
- [x] Bezeichnung, Kürzel, Besetzungshinweis speichern → `adminRoles` / `administration`
- [x] Personen hinzufügen, bearbeiten, entfernen → `admin`
- [x] Mehrfachauswahl aus Microsoft 365
- [x] Neue Rolle unter Schulleitung/Verwaltung
- [x] Rolle löschen (nur ohne Personen); Layout-Knoten entfernen
- [x] Farben weiter nach tatsächlicher Besetzung

---

## 11. Phase-3-Abnahme

- [x] Doppelklick / Stift auf ARGE-Karte → Pop-up
- [x] Bezeichnung, Kürzel, Fächer speichern → `arges`
- [x] Leitung: Name/E-Mail; Auswahl aus Lehrerliste; Übernahme aus Microsoft 365; leeren
- [x] Lehrer-Bezug: Treffer in `teachers` wird in Meta/Hinweis angezeigt
- [x] Neue ARGE unter „ARGEs / Fachschaften“
- [x] ARGE löschen (+ Layout-Knoten)
- [x] Farben: Leitung offen = vacant, besetzt = single

---

## 12. Phase-4-Abnahme

- [x] Doppelklick / Stift auf Klassen-Karte → Pop-up
- [x] Bezeichnung, Kürzel, Abschlussjahr, KV → `classes`
- [x] KV: Lehrerliste, Microsoft 365, leeren; Lehrer-Bezug in Meta
- [x] Neue Klasse unter „Klassen“
- [x] Schülervertretung: Name/E-Mail/Klasse → `studentCouncil`
- [x] SGA: Name/E-Mail/Bereich → `sga`
- [x] Neu anlegen unter „Schülervertretung & SGA“; Löschen inkl. Layout-Knoten
- [x] Farben: KV/E-Mail offen = vacant, gesetzt = single

---

## 13. Phase-5-Abnahme

- [x] `convertVerwaltungOrganigrammToSchul` mappt group/role/sammel → section/role/custom
- [x] Auto-Migration wenn `schulOrganigramm` leer und Legacy vorhanden
- [x] Sync ergänzt ARGEs/Klassen/… nach Migration
- [x] Toolbar „Aus Verwaltung übernehmen“ (erneuter Import mit Bestätigung)
- [x] Verwaltung-Tab bindet Schul-Organigramm ein (kein aktiver Builder mehr)
- [x] Legacy `verwaltungOrganigramm` bleibt archiviert (Preserve bei Teil-Saves)

---

## 14. Phase-6-Abnahme (erster Abschluss)

- [x] Textsuche über Label / Kürzel / Meta (inkl. Bereichsfilter)
- [x] Drucken-Button + Print-CSS (Chrome/Toolbar ausgeblendet)
- [x] DnD: Karten in/zwischen erlaubten Bereichen sortieren; Bereiche umordnen
- [x] IT-Sicherungs-Abgleich: Schul-Organigramm (Knoten + Einträge), Legacy klar benannt
- [x] Meta-Zeile zeigt Eintragszahl (`content`)

---

## 15. Bugfix: Sync-Vervielfachung

Ursache: `core.schulOrganigramm` wurde fälschlich mit der **Verwaltungs**-Organigramm-Normalisierung gespeichert → `refKind`/`refKey` gingen verloren → Sync legte dieselben Rollen/ARGEs bei jedem Speichern erneut an.

Fix: eigene `normalizeCoreSchulOrganigramm` + Dedupe/Repair beim Laden. Einmal Organigramm öffnen speichert die bereinigte Struktur.

---

## 16. Nächste Verbesserungen (nach Abschluss)

1. Lehrerliste: eigener Section-Zweig oder nur Verweise
2. DnD-Feinschliff (Touch, Tastatur, Undo)
3. Optional: Legacy-Feld `verwaltungOrganigramm` bereinigen / Export-Hinweis
4. Drucklayout Feinschliff (Kopfzeile mit Schulname)

---

## Tests

```bash
npm test -- tests/schul-organigramm-logic.test.mjs tests/schul-organigramm-edit.test.mjs tests/verwaltung-page.test.mjs
```
