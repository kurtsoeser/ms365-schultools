# Stammdaten Lokal-First (Umbau)

**Status:** aktiv (`STAMMDATEN_LOCAL_FIRST = true` in [`src/shared/stammdaten-local-first.js`](../src/shared/stammdaten-local-first.js))  
**Verwandt:** [schul-struktur-kanon.md](schul-struktur-kanon.md), [stammdaten-v1-sunset.md](stammdaten-v1-sunset.md)

## Ziel

Eine Schreibwahrheit lokal im Browser, ohne Sync-Chaos:

| Schicht | Inhalt |
|---------|--------|
| **Listen** (`core.*`, `years.*`) | Lehrer, ARGEs, Klassen, Schüler, Belegung, … |
| **Landkarte** (`core.schulOrganigramm`) | Hierarchie, Position, Referenzen, M365-Bindung |

Jeder Speichern-/Import-Pfad schreibt beides über [`stammdaten-write.js`](../src/shared/stammdaten-write.js) (`writeSchoolData`) und den Materializer – **kein** manueller Abgleich Listen ↔ Landkarte im Alltag.

## Was während des Umbaus aus ist

- Auto-Push/Pull SharePoint IT-Bibliothek
- Manuelles „Nach SharePoint / Von SharePoint“ (UI ausgeblendet/deaktiviert)
- Intranet-Stammdaten-Listen-Export
- IT-Bibliothek-Einrichtung im Schulregister (Panel versteckt)

## Was erlaubt ist

- **Speichern** lokal (`ms365-schooltool-data-v2`)
- **Browser-Backup** Export/Import (JSON) – manuelle Sicherung
- Landkarte und Listen-Tabs pflegen
- Importe (WebUntis / CSV) → schreiben Listen + materialisieren Landkarte

## Flag wieder einschalten (Phase 4)

1. In `stammdaten-local-first.js`: `STAMMDATEN_LOCAL_FIRST = false`
2. Zuerst nur **manueller** Upload in die IT-Bib testen
3. Auto-Sync erst danach wieder freigeben

## Phase 4 – IT-Bibliothek (noch nicht aktiv)

Geplante Reihenfolge nach stabilem lokalem Modell:

1. Manueller JSON-Export/Import (bereits Phase 0)
2. Optional manueller Upload eines v2-Snapshots in die IT-Bib
3. Auto-Sync zuletzt, Semantik: ganzer Snapshot, kein stiller Teil-Merge

Kein Multi-Browser-Versprechen, solange Auto-Sync aus bleibt.

## Entwickler-Hinweise

- Nicht auf `ms365-tenant-settings-v1` schreiben (Sunset).
- Tool-Nebenstates: [`stammdaten-side-state-keys.js`](../src/shared/stammdaten-side-state-keys.js) – Cache, nicht Wahrheit.
- `ms365TenantSettingsSave` → v2; Materializer hängt an `setCoreFromTenantSettings`.
