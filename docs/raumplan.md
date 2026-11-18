# Raumplan

**Familie:** `betrieb`  
**SoR:** SharePoint-Liste `Raeume` (Live Sync)  
**Lokal:** Plan-Bild/PDF (Data-URL / Dateiname) – kein SPO-SoR

## Rollen

| Schicht | Wert |
|---------|------|
| Portal-Persona | Lehrkraft (Audience), IT für Listen-Setup |
| App-Rolle | Lehrer: lesen/suchen + Technik-Deep-Link; Betrieb-Admin: Import, CRUD, Plan-Upload, System |

## Domäne

- Raumstammliste (Code, Name, Gebäude, Stockwerk, Kapazität, Typ, Notiz)
- CSV-Import/Export
- Plan-Viewer (Bild lokal)
- Deep-Link Technik: `technik-meldungen.html?raum=CODE`

## SPO Live Sync

- Site: `resolveSchoolToolSiteUrl('raumplan')` / `ms365-raumplan-sp-settings-v1`
- Module: `raumplan-graph.js`, `raumplan-sync.js`
- Pull beim Start; Push bei Speichern/Import/Löschen
- Demo-Seeds (`seedTag`) werden nicht gepusht; 403 wipe’t den Cache nicht
- IT-Liste: `sharepoint-liste-raumplan.html`

## Smoke

1. Lehrer: Übersicht/Suche/Technik-Link  
2. Admin: Raum speichern → Item in SharePoint sichtbar (`spoItemId`)  
3. Demo laden bleibt lokal; Pull ersetzt durch Live-Liste ohne Wipe bei 403  
