# Datenlandkarte

**Parallel zur Schuldaten-Karte (Graph):** hier geht es nicht um einzelne Personen, sondern um **Listen und Datenquellen** in der App als **Blöcke** mit **Anzahl** und **Beschreibung**. Linien = dokumentierte **logische Verknüpfung** (Sync, Join über Codes, Felder).

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
| SharePoint (Microsoft Graph `items/$count`, ConsistencyLevel eventual) | **Zeilen** – nach Anmeldung (Sites.Read.All) |

Site-URLs: Intranet aus Einrichtung; Planer-Sites aus localStorage (`ms365-sa-site-url`, `ms365-pw-site-url`, `ms365-freistellung-planer-site-v1`, `ms365-akt-planer-site-v1`) mit Fallback auf Intranet.

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
