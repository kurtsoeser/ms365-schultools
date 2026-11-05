# Schulstruktur-Kanon (Landkarte → Listen → M365 → Import)

**Status:** Zielbild / Leitplanke (Oktober 2026)  
**Produktentscheidung:** Die **Schul-Landkarte** ist die visuelle Wahrheit der Schule; die **Listen im v2-Container** bleiben die speicherbare Wahrheit. Beides muss immer zusammenpassen – inkl. Backup, WebUntis und später Bildungsportal.

**Umbau aktiv:** [stammdaten-local-first.md](stammdaten-local-first.md) – SharePoint-IT/Intranet-Export pausiert; Schreibpfad `writeSchoolData` + Materializer (kein Abgleich-Schritt im Alltag).  
**Inventar Speicherorte:** [daten-inventar-landkarte.md](daten-inventar-landkarte.md).

Verwandt: [schul-organigramm.md](schul-organigramm.md), [schul-organigramm-landkarte.md](schul-organigramm-landkarte.md), [datenlandkarte.md](datenlandkarte.md), [edtech-bildungsportal-stammdatenanbindung.md](edtech-bildungsportal-stammdatenanbindung.md)

---

## 1. Warum das zählt

Ziel ist nicht „noch ein Organigramm“, sondern:

> Die ganze Schule so abbilden, wie sie in der Microsoft-365-Welt leben soll – und alle Importe/Abgleiche schreiben genau in dieses Modell.

Praktisch:

1. Auf der Landkarte Struktur anlegen (Bereiche, Jahrgänge, Klassen, ARGEs, …).
2. Pro Bereich M365-Objekte **matchen oder anlegen** (Gruppe, SG, MESG, Verteiler, Postfach).
3. WebUntis / Bildungsportal füllen dieselben Listen und denselben Baum – ohne Parallelwelt.

---

## 2. Eine Kanon-Schicht, drei Sichten

| Sicht | Aufgabe | Schreibrecht |
|-------|---------|--------------|
| **Listen** (`core.*`, `years.*`) | Personen, Klassen, ARGEs, Belegung, … | Import & Editoren |
| **Landkarte** (`core.schulOrganigramm`) | Hierarchie, Position, M365-Bindung pro Knoten | Organigramm-UI |
| **Tenant/M365** (Graph) | Gruppen, Postfächer, Teams | Match/Create aus Area-Editor + Fachtools |

```text
Bildungsportal (Personen/Klassen)     WebUntis (Unterricht/Kursteams)
              │                                    │
              ▼                                    ▼
        years.*.students/classes            years.*.unterrichtsbelegung
        core.teachers                       → Kursteam-Knoten unter Klassen
              │
              ▼
     core.schulOrganigramm  ◄── Layout + m365 + externalIds
              │
              ▼
     Browser-Backup / IT-Sicherung  (= ganzer v2-Container)
```

**Regel:** Kein Organigramm-only-Zustand ohne Listen-Abbildung. Knoten **referenzieren** Listen (`refKind` + `refKey`). M365-Bindung liegt am Knoten (`m365`), nicht doppelt in den Personenlisten.

---

## 3. Datenmodell (Erweiterung)

### 3.1 Bereits kanonisch

| Feld | Inhalt |
|------|--------|
| `core.schulOrganigramm` | Areas, Leaves, `view`, Nesting, Kursteams unter Klassen |
| `node.m365` | `groupMode`, **`groupKind`**, `groupId`/`groupMail`, Postfach-Match |
| `years.*.unterrichtsbelegung` | Kursteams → Platzierung unter Klasse(n) |
| `setup.matched` / `catalogLinks` | Standardgruppen Lehrer/Schüler/ARGE/Klasse |

### 3.2 Neu / verbindlich für Importe

Pro Knoten optional:

```text
externalIds?: {
  bip?: string,          // Bildungsportal / Datenverbund
  sokrates?: string,     // falls aus BIP mitgeliefert
  webuntis?: string,     // Untis-Objekt-ID
  graphGroupId?: string  // Spiegel zu m365.groupId (optional)
}
```

Pro Person in Listen (bereits teilweise vorhanden):

| Liste | Stabiler Join |
|-------|----------------|
| `students[]` | `email`, optional `externalId` / BIP |
| `teachers[]` | `email`, `code` |
| `classes[]` | `code`, optional externes Klassen-ID |
| Belegung | `klasse` × `fach` × Lehrkraft |

### 3.3 Funktions-Sammelbereiche (Regel)

**Regel:** Wird dieselbe Funktion in mehreren Gruppen/Teams besetzt (Klasse, ARGE, …), legt die Landkarte **automatisch einen Sammelbereich** an, der alle Inhaber:innen dieser Funktion schulweit bündelt. An diesem Bereich hängen wieder Microsoft-365-Objekte (Gruppenchat, M365-Gruppe, Verteiler, Sicherheitsgruppe, freigegebenes Postfach).

```text
Funktion in Gruppe A / B / C  →  Stammdaten-Sitz
                              →  Sammel-Area (View)
                              →  node.m365 (Chat / Gruppe / Postfach / …)
```

Bereits umgesetzt (gleiche Mechanik):

| Funktion | Sitz-Wahrheit | Sammelbereich |
|----------|---------------|---------------|
| Klassenvorstand | `classes[].headEmail` | Root `sec_klassenvorstand` |
| ARGE-Leitung | `arges[].headEmail` | unter ARGEs `sec_arge_leiter` |
| Klassensprecher | `classes[].functions.klassensprecher` | Root `sec_classfn_klassensprecher` |

**Registry** (`listSchulOrgFunctionAggregates`): Builtin + optionale `settings.classFunctions` / `settings.argeFunctions`. Jede neue Funktions-Definition mit `aggregate: true` (Default) erzeugt beim Sync ihren Sammelbereich; M365 bleibt am Knoten (`keepM365`).

Klassenfunktionen (Beispiel):

```text
years.*.classes[].functions?: {
  klassensprecher?: { name: string, email: string }
}
```

ARGE-Funktionen analog: `core.arges[].functions` + Sammelbereich `sec_argefn_{id}` unter ARGEs.

**Verhalten am Gruppenknoten**

1. Sitz in der Gruppe zuweisen/ändern (z. B. Klasse → Klassensprecher).
2. Sammelbereich aktualisiert sich automatisch (kein manuelles Kopieren).
3. Am Sammelbereich: Microsoft-365-Objekte matchen/anlegen wie bei jeder anderen Area.

### 3.4 Zuordnungsregeln (Import)

| Quelle | Ziel |
|--------|------|
| BIP `readorgdata` Klassen | `years.*.classes` → Sync → `type:class` unter Jahrgang |
| BIP `readuserdata_v3` Schüler | `years.*.students` mit `klasse` → erscheinen im Klassen-Zoom |
| BIP Lehrkräfte | `core.teachers` |
| WebUntis Lessons / Kursteams | `unterrichtsbelegung` → `type:kursteam` unter **jeder** verknüpften Klasse |
| Manuelle Areas | bleiben; Import **merged**, löscht Custom-Bereiche nicht |

Mehrklassen-Kursteams: ein Team, **mehrere** Knoten (eine Platzierung je Klasse) – gleiche `meta.stableKey` / Graph-ID.

---

## 4. Persistenz & Sicherung

| Kanal | Inhalt |
|-------|--------|
| `ms365-schooltool-data-v2` | Kanon (inkl. `schulOrganigramm` + `m365.groupKind`) |
| `ms365-tenant-settings-v1` | Spiegel / Komfort |
| Browser-Backup | ganzer v2-Container |
| IT-Sicherungsbibliothek | ganzer v2-Container |
| SharePoint-Stammdaten-Sync | Listen; Organigramm-Layout mit Core mitwandern |

**Preserve:** Teil-Saves dürfen `schulOrganigramm` nie stillschweigend leeren.

---

## 5. Phasen-Roadmap

| Phase | Inhalt | Ergebnis |
|-------|--------|----------|
| **S0 – Leitplanke** | Dieses Dokument; Schema `groupKind` + Docs | Orientierung |
| **S1 – Persistenz hart** | Normalize/Backup/IT-Bib: `m365.groupKind`, `externalIds` | keine Datenverluste |
| **S2 – Import-Join** | Nach Belegung/`setCoreFromTenantSettings` → `syncSchulOrganigrammFromAppData` | **umgesetzt** (`schul-organigramm-stammdaten-sync.js`) |
| **S3 – Bildungsportal** | BIP-Knopf → Preview → Write in Listen → `syncSchulOrganigramm` | Personen/Klassen richtig |
| **S4 – Provisioning-Batch** | „Alle offenen Areas in M365 anlegen“ (Queue) | **umgesetzt** (`schul-organigramm-provision-diff.js` + Panel „Abgleich“) |
| **S5 – Abgleich-UI** | Diff Landkarte ↔ Listen ↔ Graph | **umgesetzt** (Struktur-Diff, Heal aus Listen, Graph-Check) |

Lehrer-/Schüler-/ARGE-/Klassen-Gruppen bleiben der **Standardpfad** (Einrichtung). Freie Areas nutzen Match/Create flexibel.

---

## 6. Nicht verwechseln

| Werkzeug | Rolle |
|----------|-------|
| **Schul-Landkarte** | Struktur + M365-Bindung pflegen |
| **Datenlandkarte** | Listen/Quellen/Sync visualisieren |
| **Schulgraph** | Beziehungen erkunden (nicht schreiben) |
| **Kursteams-Tool** | Unterrichtsteams aus WebUntis bauen |
| **BIP-Anbindung** | Personen/Klassen aus dem Datenverbund |

---

## 7. Nächster konkreter Schritt

1. **S3 – Bildungsportal:** BIP → Preview → Listen → `syncSchulOrganigramm`.  
2. Areas im Editor auf „Neu anlegen“ setzen (oder im Abgleich-Panel vormerken) → Batch über **Abgleich**.  
3. Nach Importen weiterhin `syncSchulOrganigrammFromAppData` / Heal im Abgleich nutzen.

**Bedienung S4/S5:** Schul-Organigramm → Button **Abgleich** → offene Anlagen wählen → „Ausgewählte anlegen“; Struktur-Diff heilen; optional Graph prüfen.

Die Perfektion ist erreichbar, wenn **eine** Kanonik gilt: Listen + Landkarte + M365-IDs im selben Container – und jeder Import nur diese Kanonik füttert.
