# Analyse 02 – Code-Optimierung Tool für Tool

**Stand:** 2026-09-27  
**Leitplanken:** `src/shared/ARCHITECTURE.md` – Soll 100–400 Zeilen/Modul, hartes Limit &lt; 600.  
**Goldene Regel:** Erst Split/Move ohne Verhaltensänderung, dann funktionale Verbesserungen.

---

## Kurzfassung

Die App hat bereits **gute Vorbilder** (Kursteams, Projektwochen, ARGE/Jahrgang ESM). Die größten Wartungsrisiken sind:

1. Shared-Monolithen: `tenant-settings-ui.js` (~5900), `setup-wizard.js` (~5200)  
2. Tool-Monolith: `schulstruktur-sync.js` (~5600) trotz Teil-Split  
3. Weitere 1000–2300-Zeilen-Dateien (Personen, Jahrgangsgruppen, OneNote, Gäste, …)  
4. Querschnitt: ~30 lokale `escapeHtml`/`normStr`, parallele Graph/MSAL-Clients  

Kein Feature soll verloren gehen – Optimierung = Lesbarkeit, Testbarkeit, weniger Duplikation.

---

## 1. Größen-Radar (Ist)

### Shared (Top)

| Datei | ~Zeilen | Status |
|-------|--------:|--------|
| `tenant-settings-ui.js` | 5898 | kritisch |
| `setup-wizard.js` | 5175 | kritisch |
| `app-data-v2.js` | 1689 | groß |
| `graph-unified-groups.js` | 1230 | groß, zentral |
| `graph-licenses.js` | 1138 | groß |
| `msal-auth-ui.js` | 1027 | groß |
| `browser-backup.js` | 984 | mittel–groß |

### Tools (Top)

| Datei | ~Zeilen | Cluster |
|-------|--------:|---------|
| `schulstruktur-sync.js` | 5558 | Kern |
| `jahrgangsgruppen.js` | 2350 | Unterricht |
| `personen-verwaltung.js` | 2258 | Personen |
| `onenote-verteilung.js` + `-graph.js` | 1798 + 1401 | Unterricht |
| `projektwochen-ui.js` | 1530 | Intranet |
| `eltern-verteiler.js` | 1453 | Kommunikation |
| `gast-zugaenge.js` / `gast-einlader.js` | ~1400 | Personen |
| `kursteam-templates.js` | 1402 | Unterricht |
| `schueler-lehrer-gruppen.js` | 1377 | Personen |
| `arge-fachgruppen.js` | 1360 | Unterricht |
| `elternsprechtag-bookings.js` | 1222 | Kommunikation |
| `klassen-umbenennen.js` | 1108 | Schuljahr |
| `verwaltung-gruppenverwaltung.js` | 1106 | Personen |
| `arge.js` / `jahrgang.js` | ~1070–1224 | bereits ESM-geteilt |

---

## 2. Tool-für-Tool: Zustand → Hebel → Plan

### Legende

- **Vorbild** – gut gesplittet / ESM  
- **Teil** – Module vorhanden, Entry noch zu groß  
- **Monolith** – 1–2 Riesen-Dateien  
- **Schlank** – unter Limit  

---

### 2.1 Schulstruktur-Sync — Teil / kritisch

| | |
|---|---|
| **Zustand** | 16 Module, Entry noch ~5558 Z. (UI, Graph, bind, Tenant-Detail, Archiv) |
| **Hebel** | Graph-Client (~Z. 1270–1912), Tenant-Detail/Archive-UI, `bind()`, Structure-Formulare |
| **Kürzung ohne Feature-Verlust** | Extract `-graph.js`, `-tenant-detail-ui.js`, `-bind.js`; Entry nur Orchestrierung |
| **Aufwand** | 3–5 Sessions |

**Plan:** Pro Session ein Extract + Smoke-Test + bestehende Unit-Tests. Keine Logikänderung.

---

### 2.2 Kursteams (+ Templates) — Vorbild

| | |
|---|---|
| **Zustand** | ~27 Files; größte Einzeldatei `kursteam-members.js` ~1053 |
| **Hebel** | Members + Steps-Export weiter splitten; Graph an Shared-Client |
| **Kürzung** | Lokale Utils → `shared/utils`; Filter schon klein |
| **Aufwand** | 1–2 Sessions Feinschliff |

---

### 2.3 Jahrgangsgruppen / Jahrgang — Monolith + Legacy-Doppel

| | |
|---|---|
| **Zustand** | Dashboard nutzt `jahrgangsgruppen.js` (~2350); Legacy `jahrgang/` ESM (5 Files) |
| **Hebel** | Konsolidieren: eine Codebasis; Match/Sync/PS1 trennen |
| **Risiko** | Zwei Naming-/Storage-Pfade (siehe Analyse 01 M8) |
| **Plan** | 1) Pure Logic aus jahrgangsgruppen extrahieren + Tests 2) UI/Graph splitten 3) langfristig Legacy-Jahrgang nur noch als Redirect/Export |

---

### 2.4 Personen-Verwaltung — Monolith

| | |
|---|---|
| **Zustand** | 1 File ~2258 Z. |
| **Split-Soll** | `-search-ui.js`, `-create-ui.js`, `-license-panel.js`, `-graph.js`, `-logic.js`, Entry |
| **Aufwand** | 2–3 Sessions |

---

### 2.5 Schüler-/Lehrer-Gruppen (SLG) — Teil

| | |
|---|---|
| **Zustand** | Entry ~1377 + `slg-live-details` ~1050 + `slg-gruppenverwaltung` ~905 |
| **Hebel** | Sync/Diff in Shared (`membership-reconcile`) härten; Entry weiter entschlacken |
| **Plan** | ESM-Migration der 3 Files; Sync-Logik testbar machen (Analyse 01 K2/K3) |

---

### 2.6 Gäste (Einlader / Zugänge / Hub) — Monolithen + dünner Hub

| | |
|---|---|
| **Zustand** | Hub ~114; Einlader/Zugänge je ~1400 |
| **Plan** | Wie ARCHITECTURE Phase 2: zuerst `gast-einlader` (klare Pure-Logik), dann Zugänge; gemeinsame Graph-Helfer |

---

### 2.7 OneNote-Verteilung — Monolith-Paar

| | |
|---|---|
| **Zustand** | UI ~1798 + Graph ~1401 |
| **Hebel** | UI: Steps / Preview / Progress / Apply; Graph behalten (429-sensibel) |
| **Plan** | UI-Split; Batch-Progress-Pattern aus SS wiederverwenden |

---

### 2.8 Projektwochen / Schularbeiten — Vorbild, UI dick

| | |
|---|---|
| **Zustand** | Viele Module; UI-Dateien &gt;1000 Z. |
| **Kürzung** | Demo-Seed lazy/Dev-only; Tabellen/Kalender-Render extrahieren |
| **Aufwand** | 1–2 Sessions je Tool |

---

### 2.9 Kommunikation (Postfächer, Verteiler, Eltern, Bookings)

| Tool | Zustand | Plan |
|------|---------|------|
| Postfächer | 2 Files, ~1336 gesamt | Graph-Load schon getrennt; UI unter 600 bringen |
| Verteilerlisten | ~849 | knapp über Limit; CSV/PS1 extract |
| Eltern-Verteiler | Monolith ~1453 | parse/state/graph/ui |
| Elternsprechtag | ~1222 + logic | Bookings-Graph vs. Setup-UI |

---

### 2.10 SharePoint / Intranet — modular, SPO-Duplikat

| | |
|---|---|
| **Zustand** | `spo-graph-shared` + Listen-Tools |
| **Hebel** | List-Scaffold (Site-Resolve, Ensure List, Columns) vereinheitlichen |
| **Plan** | Ein `spo-list-kit.js`; Listen-Tools werden dünne Konfiguration + Schema |

---

### 2.11 ARGE / Fachgruppen / Verwaltung / Klassenvorstände

| Tool | Zustand | Plan |
|------|---------|------|
| `arge/` | ESM-Vorbild | Shared „group provision“ mit Jahrgang |
| `arge-fachgruppen` | Monolith ~1360 | an ARGE-Pattern angleichen |
| Verwaltung | ~1106 | Gruppenverwaltung splitten |
| Klassenvorstände | ~786 | leicht über Limit; State/Graph extract |

---

### 2.12 Organisations-Assistent — mittel, gut testbare Logic

| | |
|---|---|
| **Zustand** | Entry + logic + cohorts |
| **Hebel** | Persistenz mit SS vereinheitlichen (Analyse 01 K4); Logic schon relativ sauber |

---

### 2.13 Lizenzverwaltung — mittel

| | |
|---|---|
| **Zustand** | ~809 + Shared Licenses |
| **Plan** | UI vs. Assign trennen; keine lokale Lizenz-Logik duplizieren |

---

### 2.14 Power Automate — schlank (Doku/Setup)

| | |
|---|---|
| **Zustand** | Catalog + Recipe-Page + viele dünne HTMLs |
| **Kürzung** | Ein Template für alle `pa-*.html`; `escapeHtml` Shared |

---

### 2.15 Weitere Teams & Gruppen (WTG) — Vorbild, Katalog-unsichtbar

| | |
|---|---|
| **Zustand** | 4 ESM-Files, Multi-Step-HTML |
| **Plan** | Graph an Shared-Client; optional eine Wizard-Shell statt 5 HTMLs |

---

### 2.16 Schlanke / Legacy-Tools

| Tool | Hinweis |
|------|---------|
| Namenskonvention, Diplomarbeiten, Datei-Migration, Leere-Gruppen, Cleanup-Playbook | unter/knapp Limit – nur Utils-Sweep |
| Teams-Archiv | Redirect/Archiv; kein neues Feature; Code ggf. entfernen |
| Klassen-Umbenennen / Merge | funktional kritisch → erst Tests (Analyse 01), dann Split |

---

### 2.17 Shared Setup / Tenant — kritisch (App-weit)

| Modul | Plan |
|-------|------|
| `tenant-settings-ui.js` | Tabs → eigene Module (Klassen, Lehrer, Schüler, Domain, Backup-UI); Core bleibt |
| `setup-wizard.js` | Steps als `-step-*.js`; Admin-Model schon teilweise da |
| `graph-unified-groups.js` | wird kanonischer Client; Tool-lokale `graphRequest` entfernen |
| `msal-auth-ui.js` | einzige PCA; Tool-PCAs abschaffen |

---

## 3. Querschnitt-Optimierungen (hoher Hebel, niedriges Risiko)

| Maßnahme | Effekt | Aufwand |
|----------|--------|---------|
| `escapeHtml` / `normStr` / `safeJsonParse` Sweep | −Duplikate, weniger Bugs | 1–2 Tage |
| Shared Graph-Client (Token, 429, Paging, Truncation-Flag) | weniger Auth-/Paging-Bugs | 3–5 Tage |
| Shared CSV parse/download | weniger Parser-Drift | 1 Tag |
| Shared Bulk-Progress UI | einheitliches Feedback | 1–2 Tage |
| ESM + `type="module"` Rest | Build/Tree-shaking | laufend |

---

## 4. Umsetzungsplan (Phasen)

### Phase A – Fundament (Woche 1–2)

1. Shared Graph-Client-API spezifizieren (Wrapper um bestehende `graph-unified-groups`).  
2. Truncation + Identity-Helfer (mit Analyse 01).  
3. `escapeHtml`-Sweep in Top-10-Dateien.  
4. Zentrale Schuljahr-Funktion.

**DoD:** Kein neues Tool-Feature; Build + Tests grün; 1 Pilot-Tool nutzt nur Shared-Client.

### Phase B – Top-Monolithen splitten (Woche 3–6)

Reihenfolge (wie ARCHITECTURE, aktualisiert):

1. `schulstruktur-sync.js` Graph + bind + tenant-detail  
2. `personen-verwaltung`  
3. `jahrgangsgruppen` (Logic zuerst wegen Risiko)  
4. `onenote-verteilung` UI  
5. `gast-einlader` → `gast-zugaenge`  
6. `eltern-verteiler` / `arge-fachgruppen`

**DoD pro Tool:** Entry &lt; 800 Z. (Zwischenziel), danach &lt; 600; Smoke-Test HTML; keine Feature-Regression.

### Phase C – Shared UI-Riesen (Woche 7–10)

1. `tenant-settings-ui` nach Tabs  
2. `setup-wizard` nach Steps  
3. Restliche IIFE → ESM

**DoD:** ESLint `max-lines` Warnungen für Tools deutlich reduziert; Shared-Warnungen dokumentiert.

### Phase D – Konsolidierung (laufend)

1. jahrgang ↔ jahrgangsgruppen  
2. arge ↔ arge-fachgruppen  
3. SPO List-Kit  
4. WTG Wizard-Shell  
5. Teams-Archiv-Code aufräumen  

---

## 5. Priorisierte „Kürzungsliste“ (ohne Feature-Verlust)

| Prio | Was | Geschätzte Zeilen-Reduktion Entry |
|-----:|-----|-----------------------------------|
| 1 | SS Graph+bind extract | −1500–2500 aus Entry |
| 2 | Personen split | 2258 → 5× &lt;500 |
| 3 | Jahrgangsgruppen Logic/UI | 2350 → testbare Logic + UI |
| 4 | OneNote UI steps | 1798 → 3–4 Module |
| 5 | escapeHtml/norm Sweep | −lokale Kopien app-weit |
| 6 | Graph-Client Konsolidierung | −hunderte Duplikat-Zeilen in 5+ Tools |

---

## 6. Was bewusst *nicht* kürzen

- Demo-Daten-Inhalte (nur lazy laden)  
- PowerShell-Export-Strings (fachlich nötig)  
- Schema-Konstanten SharePoint/sRDP  
- Bereits gut getestete Pure-Logic in Kursteams/SS-Match  

---

## 7. Checkliste pro Refactor-PR

- [ ] Nur ein Tool oder ein Shared-Modul  
- [ ] Move first, keine Verhaltensänderung  
- [ ] `npm test` + `npm run build`  
- [ ] HTML `type="module"` falls ESM  
- [ ] Smoke: Login → Hauptaktion → Speichern  
- [ ] Zeilen Entry dokumentieren (vorher/nachher)
