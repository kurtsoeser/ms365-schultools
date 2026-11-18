# Ausgabenverwaltung (Schlüssel, Karten, Zugang)

**Stand:** 2026-11-18  
**Status:** P5  
**Familie:** `betrieb`  
**Roadmap:** [`schul-apps-roadmap-prozesse.md`](./schul-apps-roadmap-prozesse.md)  
**Lizenzpaket:** `ausgaben`  
**SPO Live Sync:** [spo-live-sync.md](./spo-live-sync.md) · Tracker [#03](./spo-live-sync-tracker.md#03-ausgaben)

---

## Idee

Eine Oberfläche: **Was wurde an wen ausgegeben?**  
Vorlagen für Schlüssel, Raumkarte, Schrankenkarte, Parkplatz, Laptop – Schulen legen **eigene Typen** an (z. B. Beamer-Koffer). SharePoint speichert Typ als Text, nicht als starre Choice-Liste.

## Funktionen

- Inventar pflegen (CSV-Import/-Export)
- Ausgeben / Zurücknehmen / verloren
- Filter: Typ, Suche, überfällig
- Übergabeprotokoll (Textdatei)
- Typen & Vorlagen im Tool (lokal; Typ-Text landet in SPO am Gegenstand)
- SharePoint-Paket: `AusgabenGegenstaende` + `AusgabenVorgaenge`

## System of Record

| Schicht | Rolle |
|---------|--------|
| SharePoint-Listen `AusgabenGegenstaende` + `AusgabenVorgaenge` | **SoR** |
| App | Inventar / Ausgeben / Retour + Pull/Push |
| localStorage | Cache, Demo (`localDemoOnly` / `seedTag`) |

Site über `resolveSchoolToolSiteUrl('ausgaben')` bzw. `ms365-ausgaben-sp-settings-v1`.  
Module: `ausgaben-graph.js`, `ausgaben-sync.js`. Demo-Seeds werden nicht nach SPO gepusht; 403 wipe’t den lokalen Stand nicht.

## Rollen (Kurz)

| Schicht | Inhalt |
|---------|--------|
| Portal-Persona | `lehrer` / `it` |
| App-Rolle Admin | Inventar, Ausgeben, Retour, Settings, SPO-Pull |
| App-Rolle Lehrkraft | „Meine Ausgaben“ (`empfaengerEmail`) |

## Nicht im Scope

Zahlungs-/Kautionseinzug, Schranken-API, Kennzeichenerkennung.

## Tool-IDs

- `ausgaben`
- `sharepoint-liste-ausgaben`
