# Kollegiumsumfragen

**Stand:** 2026-11-18  
**Status:** P4 (MSAL + SharePoint SoR / Live Sync)  
**Roadmap:** [`schul-apps-roadmap-prozesse.md`](./schul-apps-roadmap-prozesse.md)  
**Lizenzpaket:** `ablauf`  
**Familie:** `ablauf-kommunikation`  
**SPO Live Sync:** [`spo-live-sync.md`](./spo-live-sync.md) · Tracker #10

---

## Warum SharePoint – nicht Forms?

Viele Schulen kennen Microsoft Forms. Für Schultools ist **SharePoint** die bessere Datenbasis:

- Listen sind erweiterbar (Spalten, Power Automate, Rechte)
- Fragen/Antworten als JSON in Listen → komplexe Umfragen ohne Forms-Lock-in
- Auswertung und Verwaltung bleiben im Schultool
- Provisionierung: ein Klick legt das Listen-Paket an

Forms kann später optional als **Zusatzkanal** (Link/Export) dienen – nicht als System of Record.

## System of Record

| Schicht | Rolle |
|---------|--------|
| SharePoint `KollegiumUmfragen` + `KollegiumUmfrageAntworten` | **SoR** |
| App | Formular, Rollen-UI, Pull/Push |
| localStorage | Cache / Offline / Demo (`localDemoOnly`, `seedTag`) |

- Site: `resolveSchoolToolSiteUrl('kollegium')` + `ms365-kollegium-umfragen-sp-settings-v1`
- Graph/Sync: `kollegium-umfragen-graph.js` · `kollegium-umfragen-sync.js`
- Start → Pull; Schreiben → lokal sofort → Push; `spoItemId` merken
- 403 wipe’t Cache/Demo nicht; Demo-Seeds werden nicht nach SPO gepusht

## Funktionen (MVP)

- Vorlagen: Ja/Nein, Klassenwunsch, SCHILF, Stundenplan-Randbedingungen, KV-Feedback, Verfügbarkeit, Beleg mit Upload
- Fragetypen: Single, Multi, Freitext, Skala, **Datei-Upload**
- Optional zusätzlicher Anhang zur gesamten Antwort
- Zielgruppen: gesamtes Kollegium, Klassenvorstände, Abteilung, manuelle Mail-Liste
- Optional Klassen-Bezug (Chips aus Stammdaten) und Empfänger-Text
- Lehrkräfte antworten nur mit Microsoft-Konto (`kollegium-umfragen-antwort.html`)
- Frist: nach `FaelligAm` keine neuen Antworten mehr
- Live-Auswertung + CSV im Admin-Tool (Anhänge lokal öffnen)
- SharePoint-Paket: `KollegiumUmfragen` + `KollegiumUmfrageAntworten`

## Rollen

| Schicht | Inhalt |
|---------|--------|
| Portal-Persona | `lehrer` · `it` |
| App-Rollen | Admin (Schulleitung/Verwaltung/Schul-IT) legt an/auswertet; Lehrkraft antwortet |
| Lizenz | Feature `kollegium-umfragen` im Paket `ablauf` |

## Tool-IDs

- `kollegium-umfragen`
- `sharepoint-liste-kollegium-umfragen`

## Smoke

1. IT: Listen-Assistent → Site + Listen anlegen  
2. Admin: Umfrage speichern/aktivieren → Item in `KollegiumUmfragen` sichtbar (`spoItemId`)  
3. Lehrkraft: Antwort-Seite → Abgabe → Item in `KollegiumUmfrageAntworten`  
4. 403-Simulation: lokaler Stand bleibt; Demo laden pusht keine Seeds  
