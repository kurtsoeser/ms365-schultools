# Analyse-Serie 2026-09 – Übersicht

Systematische Durchsicht der App **MS365schule** (alle Tools, Shared-Module, Dashboard).  
Ziel: Korrektheit, Wartbarkeit, Bedienung und sinnvolle Ausbaustufen für  
„Microsoft 365 für die Schule. Einfach. Alles.“

| # | Dokument | Inhalt |
|---|----------|--------|
| 01 | [analyse-01-fehler-bugs-luecken.md](./analyse-01-fehler-bugs-luecken.md) | Bugs, Matching-/Sync-Risiken, Testlücken + Umsetzungsplan |
| 02 | [analyse-02-code-optimierung-tools.md](./analyse-02-code-optimierung-tools.md) | Tool-für-Tool Größen, Splits, Duplikate + Umsetzungsplan |
| 03 | [analyse-03-ui-ux-verbesserung.md](./analyse-03-ui-ux-verbesserung.md) | Bedienung, Playbooks, Kontext, Bulk-Feedback + Umsetzungsplan |
| 04 | [analyse-04-ausbau-neue-tools.md](./analyse-04-ausbau-neue-tools.md) | Fehlende Schul-Bausteine, neue Tools, Roadmap „nächste Tage“ |

## Empfohlene Gesamt-Reihenfolge

1. **Korrektheit** (01 Sprint 1) – Match, Sync-Identität, Truncation, Schuljahr, State  
2. **UX-Kontext + Playbook Schuljahresstart** (03 + 04) – sofort spürbar  
3. **Shared Graph-Client + Monolith-Splits** (02) – parallel, risikarm  
4. **Lifecycle + Sync-Monitor** (04) – neue Substanz  

Ältere Ideen: [projektanalyse-und-werkzeugideen.md](./projektanalyse-und-werkzeugideen.md) · Architektur: [../src/shared/ARCHITECTURE.md](../src/shared/ARCHITECTURE.md)

**Methode:** Code-Review der Kernpfade + parallele Cluster-Analyse; kein Live-Tenant-Test.
