# Daten-Inventar-Landkarte

**Zweck:** Übersicht, **wo** welche Schul-/App-Daten physisch liegen – nicht die Live-Zähler der App.

| Werkzeug / Doc | Aufgabe |
|----------------|---------|
| **Diese Seite** | Inventar: Speicherorte, Wahrheit vs. Cache, was während Lokal-First aktiv ist |
| [Datenlandkarte](datenlandkarte.md) (App: `tools/datenlandkarte.html`) | Visuelle Blöcke mit **Anzahlen** und Sync-Linien zwischen Listen |
| [Schul-Landkarte / Organigramm](schul-organigramm-landkarte.md) | **Pflege** der Schulstruktur (schreibt Stammdaten) |
| [Lokal-First](stammdaten-local-first.md) | Umbau-Status: SPO/Intranet pausiert |

---

## 1. Grobbild (Lokal-First, Stand Umbau)

```mermaid
flowchart TB
  subgraph browser [Dieser_Browser]
    V2["Kanon: ms365-schooltool-data-v2"]
    Side["Nebenstates: UI_Graph_Cache"]
    UI["UI_Prefs: Theme_Layout"]
    Sess["sessionStorage: Auth_Drafts"]
  end

  subgraph manuell [Manuell_erlaubt]
    JSON["Browser-Backup JSON Export_Import"]
  end

  subgraph pausiert [Pausiert_Lokal_First]
    IT["SharePoint IT-Bibliothek"]
    Intra["Intranet Stammdaten-Listen"]
  end

  subgraph live [Live_nicht_persistiert]
    Graph["Microsoft Graph Gruppen_Postfaecher"]
  end

  V2 --> JSON
  Side -.-> V2
  V2 -.->|spaeter_Phase4| IT
  V2 -.->|spaeter| Intra
  V2 -.->|Match_Create| Graph
```

**Merksatz:** Alles Wichtige für Schule und Landkarte steckt in **einem** localStorage-Schlüssel (`ms365-schooltool-data-v2`). Alles andere ist Cache, Site-URL, oder später wieder SPO.

---

## 2. Kanonischer Container (Wahrheit)

Schlüssel: `ms365-schooltool-data-v2`  
Code: [`app-data-v2.js`](../src/shared/app-data-v2.js), Schreiben: [`stammdaten-write.js`](../src/shared/stammdaten-write.js)

```mermaid
flowchart LR
  subgraph v2 [ms365-schooltool-data-v2]
    core["core schulweit"]
    years["years pro Schuljahr"]
    setup["setup Einrichtung"]
    structure["structure match tenant Cache"]
  end

  subgraph coreDetail [core]
    teachers[teachers]
    arges[arges]
    subjects[subjects]
    admin[administration_Rollen]
    org["schulOrganigramm Landkarte"]
  end

  subgraph yearDetail [years.byLabel]
    classes[classes]
    students[students]
    guardians[guardians]
    belegung[unterrichtsbelegung]
  end

  core --> coreDetail
  years --> yearDetail
```

| Bereich | Inhalt | Rolle |
|---------|--------|--------|
| `core.teachers` / `arges` / `subjects` / … | Listen | Entitäten-Wahrheit |
| `core.schulOrganigramm` | Baum, Position, `refKind`/`refKey`, M365-Bindung | Struktur-Wahrheit |
| `years.*.classes` / `students` / … | schuljahrbezogen | Entitäten-Wahrheit |
| `years.*.unterrichtsbelegung` | Kursteams-Zuordnung | Entitäten-Wahrheit |
| `setup.*` | Intranet-URL, Matches, Import-Historie, IT-Meta | Konfiguration |
| `structure` / `match` / `tenant.cache` | Schulstruktur-Sync | **Cache** im Container |

Legacy-Schlüssel `ms365-tenant-settings-v1`: **nur noch lesen/migrieren**, nicht schreiben ([Sunset](stammdaten-v1-sunset.md)).

---

## 3. Was lokal noch daneben liegt (kein Stammdaten-Kanon)

Katalog im Code: [`stammdaten-side-state-keys.js`](../src/shared/stammdaten-side-state-keys.js)  
Vollständiges Backup-Inventar: [`browser-backup.js`](../src/shared/browser-backup.js) (`STORAGE_CATALOG`)

```mermaid
flowchart TB
  subgraph wahrheit [Wahrheit]
    V2[ms365-schooltool-data-v2]
  end

  subgraph graphWork [Graph_Arbeitsstand]
    Arge[ms365-arge-state-v2]
    Jg[ms365-jahrgang-state-v1]
    Wtg[ms365-wtg-state-v1]
  end

  subgraph uiCache [UI_Wizard_Cache]
    Kurs[webuntis-teams-creator-state-v1]
    Strukt[ms365-schulstruktur-sync-v1]
    OrgView[ms365-schul-org-view-mode-v1]
    DlLayout[ms365-datenlandkarte-layout-v1]
  end

  subgraph sites [Planer_Site_URLs]
    SA[ms365-sa-site-url]
    PW[ms365-pw-site-url]
    FR[ms365-freistellung-planer-site-v1]
    AK[ms365-akt-planer-site-v1]
  end

  Arge -.->|ARGEs kommen aus| V2
  Kurs -.->|Belegung kommt aus| V2
  Strukt -.->|Spiegel auch in v2.structure| V2
```

| Typ | Beispiele | Darf Stammdaten ersetzen? |
|-----|-----------|---------------------------|
| Graph-Arbeitsstand | ARGE-/Jahrgang-/WTG-State | Nein |
| UI-/Wizard-Cache | Kursteams-Creator, Organigramm-Ansicht | Nein |
| Planer-Site-Keys | `ms365-sa-site-url`, … | Nein (nur Site-Ziel) |
| Theme / Dashboard | `ms365-theme-v1`, … | Nein |

---

## 4. Fluss beim Speichern (Lokal-First)

```mermaid
sequenceDiagram
  participant UI as Listen_oder_Landkarte
  participant Write as writeSchoolData
  participant V2 as localStorage_v2
  participant Mat as Materializer
  participant SPO as SharePoint

  UI->>Write: Speichern_oder_Import
  Write->>V2: Listen_plus_Layout
  Write->>Mat: fehlende_Knoten_nachziehen
  Mat->>V2: schulOrganigramm_aktualisieren
  Note over SPO: pausiert_kein_Auto_Push
```

Sicherung heute: **Browser-Backup → Exportieren** (alle `ms365-*` / `webuntis-*` Keys als JSON).

---

## 5. Unterschied zur App-Datenlandkarte

Die **App-Datenlandkarte** zeigt z. B. „12 Klassen · 340 Schüler“ und Linien „Schüler → Klasse“. Sie erklärt nicht, dass Klassen in `years.byLabel` und die Hierarchie in `schulOrganigramm` stecken.

Diese **Inventar-Landkarte** erklärt die Speicherorte und die Regel „ein Schreibpfad“. Beide ergänzen sich:

1. Entwickler / Architektur → dieses Doc  
2. Schule / Überblick Zahlen → `tools/datenlandkarte.html`  
3. Struktur pflegen → Schul-Organigramm / Landkarte  

---

## 6. Kurzes Inventar (Checkliste)

**Aktiv & lokal (Wahrheit)**

- [x] `ms365-schooltool-data-v2` – Core, Years, Organigramm, Setup  
- [x] Manuelles Browser-Backup (JSON)

**Lokal, aber nur Cache / Arbeitsstand**

- [x] ARGE / Jahrgang / WTG / Kursteams-State-Keys  
- [x] Schulstruktur-Sync-Keys  
- [x] Planer-Site-URLs, UI-Layouts

**Pausiert (Lokal-First)**

- [ ] IT-Bibliothek Auto-Sync / Upload  
- [ ] Intranet-Stammdaten-Listen-Export  

**Nicht lokal gespeichert**

- [ ] Graph-Mitgliederlisten (live)  
- [ ] MSAL-Tokens (session, bewusst nicht im Backup)
