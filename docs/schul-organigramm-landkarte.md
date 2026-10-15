# Schul-Organigramm → Schul-Landkarte (Plan)

**Zielbild:** Verschachtelbare **Bereiche** (Areas) auf einem **Zoom-/Pan-Canvas** (wie die [Datenlandkarte](datenlandkarte.md)), in denen Rollen, Personen, ARGEs, Klassen und SV/SGA-Boards liegen. Pro Bereich später: **M365-Gruppe** und/oder **freigegebenes Postfach** verknüpfen oder anlegen. **Keine Überlappungen** in Phase 1 (reine Baum-Hierarchie / Containment).

**Abgrenzung**

| Werkzeug | Aufgabe |
|----------|---------|
| **Schul-Landkarte** (dieses) | Schulorganisation pflegen + grafisch abbilden |
| [Datenlandkarte](datenlandkarte.md) | Listen/Quellen/Sync (nicht Personenstruktur) |
| [Schulgraph](schulgraph.md) | Beziehungsgraph erkunden |

---

## 1. Mentales Modell

```text
Schulverwaltung                    ← Area (kann M365-Gruppe + Postfach haben)
├── Schulleitung                   ← Area
│   ├── Direktion                  ← Role → adminRole + Personen
│   └── Stellv. Schulleitung       ← Role
├── Verwaltung                     ← Area
│   ├── Sekretariat                ← Area (mehrere User direkt oder via Rollen)
│   ├── Administration             ← Area
│   └── … weitere User an Area
└── ARGEs / Fachleitungen          ← Area
    ├── FS Mathe                   ← Arge-Ref
    └── FS Deutsch                 ← Arge-Ref
```

- **Area** = Organisationskasten (verschachtelbar).
- **Leaf-Knoten** = Rolle, ARGE, Klasse, SV-Board, SGA-Board, optional Person-Slot.
- Personen bleiben in Stammdaten-Listen (`admin`, `studentCouncil`, …); Karten **referenzieren**.

---

## 2. Persistenz (lokal + IT-Sicherung)

### 2.1 Kanonisch

Weiterhin **`ms365-schooltool-data-v2` → `core.schulOrganigramm`**.

- Browser-Backup und IT-Sicherungsbibliothek speichern den **ganzen** v2-Container → **kein Extra-Key** nötig.
- Spiegel `ms365-tenant-settings-v1` wie bisher.
- Preserve-Semantik bei Teil-Saves beibehalten (`tenant-settings-core` / `app-data-v2`).

### 2.2 Schema `version: 2`

```text
schulOrganigramm = {
  version: 2,
  view: { x: number, y: number, k: number },   // Canvas-Viewport (mitgesichert)
  nodes: OrgNode[]
}

OrgNode = {
  id: string,
  type: 'area' | 'role' | 'arge' | 'class' | 'council' | 'sga' | 'person' | 'custom',
  label: string,
  parentId: string | null,     // Verschachtelung (Baum, keine Überlappung)
  order: number,               // Geschwister-Reihenfolge
  x: number, y: number,        // Position relativ zum Parent-Inhalt (Canvas)
  w?: number, h?: number,      // optional; Auto-Layout wenn fehlend
  collapsed?: boolean,
  tone?: string,
  refKind: 'none' | 'adminRole' | 'arge' | 'class' | 'kursteam' | 'studentCouncil' | 'sga' | 'audience' | 'adminPerson' | 'jahrgang' | 'sammelgruppe',
  refKey: string,
  externalIds?: { bip?: string, sokrates?: string, webuntis?: string, graphGroupId?: string },
  m365?: {
    groupMode: 'none' | 'match' | 'create',
    groupKind: 'unified' | 'security' | 'mailSecurity' | 'distribution',
    groupId: string | null,
    groupLabel?: string | null,
    groupMail?: string | null,
    mailboxMode: 'none' | 'match',
    mailboxUpn: string | null,
    mailboxLabel?: string | null,
    mailboxId?: string | null
  }
}

**Kanon / Importe:** [schul-struktur-kanon.md](schul-struktur-kanon.md)
```

**Hinweise**

| Feld | Sync IT-Bib | Bemerkung |
|------|-------------|-----------|
| `nodes` (Baum + x/y + refs + m365) | ja | Strukturwahrheit |
| `view` (Pan/Zoom) | ja | damit Wiederherstellen „wie verlassen“ möglich ist |
| reine Browser-UI-Prefs (Theme) | nein | bleiben localStorage |

Optional später: schwere Viewport-Prefs nur lokal – **Empfehlung: `view` im Core**, analog zur Idee „Anordnung gehört zur Schule“.

### 2.3 Migration `v1 → v2`

| v1 | v2 |
|----|----|
| `type: 'section'` | `type: 'area'` |
| `parentId: null` (flach) | unverändert zuerst; Nesting manuell/UI |
| fehlende `x/y` | Auto-Grid aus `order` |
| `refKind: 'audience'` an Areas | bleibt (Verwaltungsgruppen) |
| SV/SGA `_board` | bleibt |
| `verwaltungAudienceGroups` | unverändert parallel |

Beim ersten Öffnen: `version` hochsetzen, Layout speichern (`dirty`). Alte Clients lesen nur `nodes` (Normalize tolerant).

### 2.4 IT-Sicherung / Abgleich

- Zähler `schulOrganigramm` / `schulOrganigrammContent` weiter über Node-Anzahl (Areas zählen als Struktur, Content = Leaves).
- Optional später: eigener Zähler `schulOrganigrammAreas`.

---

## 3. UI-Konzept (Canvas)

Anlehnung an Datenlandkarte:

- **Mausrad** = Zoom, **Leerfläche ziehen** = Pan, **Kopf einer Area** = Area verschieben (Geschwister / später Re-Parent mit Modifier).
- **Doppelklick / Stift** = bearbeiten (wie heute).
- **Drop in Area** = Kind werden (Nesting) oder Person einer Area zuordnen.
- **Keine Kanten-Überlappungslogik** in Phase 1: Containment-Rechtecke, Auto-Pack der Kinder.

Chrome oben: kompakt (Filter, Suche, + Bereich, Druck/Export später).

Zwei Modi (Übergang):

1. **Liste** (heutiges UI) – stabil  
2. **Landkarte** (Canvas) – neu, Default sobald stabil  

---

## 4. Phasen

| Phase | Inhalt | Ergebnis |
|-------|--------|----------|
| **A – Spez & Schema** | Spec, `version: 2`, Normalize, Migration, Tests | Daten sicher in Core + IT-Bib |
| **B – Canvas-Grundlage** | Zoom/Pan, Areas als Boxen (flach migriert), Position speichern | Grafische Ansicht nutzbar |
| **C – Nesting** | Area in Area, Re-Parent, Auto-Layout Kinder | „Schulverwaltung → …“ möglich |
| **D – Personen an Areas** | Direkte User-Slots (`refKind: adminPerson` / Membership) | **umgesetzt** |
| **E – M365-Anbindung** | Pro Area: M365-Gruppe / SG / MESG / Verteiler matchen oder anlegen; Postfach matchen | **umgesetzt** (Create via Graph; Postfach-Create bleibt Exchange-Tool) |
| **F – Feinschliff** | Druck/Export Canvas, Tastatur, Mobile | teilweise: Stift-Popup + Drill-down/Mitglieder |

**Nicht in Phase A–C:** Überlappende Areas, freie Graph-Kanten, Live-Teams-Kanäle.

---

## 5. Dateien (geplant)

| Pfad | Rolle |
|------|--------|
| `docs/schul-organigramm-landkarte.md` | Diese Spec |
| `docs/schul-organigramm.md` | Verweis + Status |
| `src/tools/schul-organigramm/schul-organigramm-logic.js` | Normalize v2, Migration, Nesting-APIs |
| `src/tools/schul-organigramm/schul-organigramm-canvas.js` | Zoom/Pan/Render (neu) |
| `src/tools/schul-organigramm/schul-organigramm-ui.js` | Modus Liste/Landkarte |
| `src/shared/app-data-v2.js` | `normalizeCoreSchulOrganigramm` tolerant v1/v2 |
| `tests/schul-organigramm-*.test.mjs` | Migration + Nesting |

Canvas-Technik: Muster aus `datenlandkarte-ui.js` (`transform: translate + scale`).

---

## 6. Abnahme Phase A+B (Minimum)

- [x] v1-Layouts migrieren ohne Datenverlust (`section`→`area`, `view`, `x/y`)
- [x] Areas/Leaves auf Canvas sichtbar, Zoom/Pan
- [x] Positionen + Viewport in `core.schulOrganigramm` → Backup/IT-Bib
- [x] Bestehende Listen-Pflege (Rollen, ARGEs, …) weiter funktionsfähig
- [x] Tests grün

---

## 7. Offene Produktentscheidungen (Defaults vorgeschlagen)

1. **Default-Ansicht:** Landkarte, Liste als Umschalter – **ja**.  
2. **Viewport in Core:** **ja** (mit IT-Bib).  
3. **Person ohne Rolle an Area:** Phase D – Membership analog `adminAudienceMemberships` oder `admin` mit Area-Key – **Empfehlung:** Area mit `refKind: 'audience'` nutzt bestehende Verwaltungsgruppe; reine Layout-Area ohne Audience erst Person-Slots.  
4. **ARGEs unter verschachtelter Area:** Seed legt Standard-Area „ARGEs“ an; Verschieben erlaubt.
