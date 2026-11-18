# Datenlandkarte

**Parallel zur Schuldaten-Karte (Graph):** hier geht es nicht um einzelne Personen, sondern um **Listen und Datenquellen** in der App als **Blöcke** mit **Anzahl** und **Beschreibung**. Linien = dokumentierte **logische Verknüpfung** (Sync, Join über Codes, Felder).

**Nicht verwechseln:**

| Doc / Tool | Aufgabe |
|------------|---------|
| **Diese Datenlandkarte (App)** | Live: Listen-Blöcke, Anzahlen, Sync-Linien |
| [Daten-Inventar-Landkarte](daten-inventar-landkarte.md) | Doku: **wo** localStorage/SPO/Graph liegen (Wahrheit vs. Cache) |
| [Schul-Organigramm](schul-organigramm.md) / [Schul-Landkarte](schul-organigramm-landkarte.md) | **Pflege** der Schulstruktur (schreibt Stammdaten) |
| [schul-struktur-kanon.md](schul-struktur-kanon.md) | Zielbild eine Kanonik |
| [Lokal-First](stammdaten-local-first.md) | Umbau: SPO pausiert |

## Öffnen

`tools/datenlandkarte.html` – Dashboard: *Einstellungen & Übersicht → Datenlandkarte*

## Blöcke (Schichten)

1. **Stammdaten** – Klassen, Lehrkräfte, Fächer, ARGE (+ Hub)
2. **Schuljahr** – Schüler:innen, Eltern, Unterrichtsbelegung, WebUntis/SIS-Import
3. **SharePoint** – einzelne Listen (Klassen, Fächer, Schülerinnen, Lehrerliste, SAP-Schularbeiten, PW-Angebote, Freistellungen, Schulaktivitäten)
4. **Planer** – Schularbeiten, Projektwochen, Freistellungen, Schulaktivitäten (Site konfiguriert ✓/—)
5. **M365** – catalogLinks (Gruppen aus Einrichtung)

### Zähler

| Quelle | Anzeige |
|--------|---------|
| Browser (`tenant-settings`, `app-data-v2`, Import-Historie) | **Einträge** |
| SharePoint (Graph `$count` → `@odata.count` → Seitenweise; sonst **✓** wenn Liste gefunden) | **Zeilen** bzw. vorhanden – nach Anmeldung (Sites.Read.All) |

Site-URLs: Intranet aus Einrichtung; Planer-Sites über Stammdaten **Planer & Apps** (`planerDefaultSiteUrl` / Overrides) bzw. Tool-localStorage, Fallback Intranet.

**Aktualisieren** lädt lokale Metriken sofort und SharePoint-Zähler asynchron. Fehler erscheinen als Statuszeile (z. B. fehlende Site, Liste nicht gefunden).

### Anordnung

- **Kopfzeile** eines Blocks (Griff-Symbol) **ziehen** → Block verschieben; Verbindungslinien passen sich live an.
- Positionen werden unter `ms365-datenlandkarte-layout-v1` in **localStorage** gespeichert (pro Browser).
- **Standardlayout** setzt die automatische Schicht-Anordnung zurück.

## Verbindungen (Beispiele)

| Von | Nach | Bedeutung |
|-----|------|-----------|
| Schüler:innen | Klassen | Feld `klasse` / Code |
| Lehrkräfte | Klassen | Klassenvorstand (`headEmail`) |
| Unterrichtsbelegung | Klasse / Lehrkraft / Fach | Kursteam-Endliste |
| SharePoint Stammdaten | Stammdaten | Export/Sync |
| Schularbeiten-Planer | Klassen, Fächer, Lehrkräfte | Join über Codes (kein SP-Lookup) |

## Erweiterungen (optional)

- Filter „nur Sync-Ketten“ / „nur Planer“
- Block für SP „PW-Aktionen“ (Zähler bereits in Graph-Probe)

## Tests

```bash
npm test -- tests/datenlandkarte-layout.test.mjs tests/datenlandkarte-spo-metrics.test.mjs
```
