# Technik-Meldungen

**Stand:** 2026-11-18  
**Familie:** `betrieb`  
**UI:** `tools/technik-meldungen.html`, `src/tools/technik-meldungen/`  
**IT-Sibling:** `sharepoint-liste-technik-meldungen` (Audience `it`)  
**Rollenvertrag:** [`rollen-vertrag.md`](./rollen-vertrag.md) Welle 4  
**Upgrade-Vorlage:** [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md)

---

## Zielbild

```text
Schüler / Lehrkraft ── meldet Problem (eigene Tickets)
Betrieb-Admin (IT/Schulleitung/Verwaltung) ── Inbox, Status-Pipeline, Auswertung
Schul-IT ── Site/Liste (Listen-Assistent), Stammdaten
                      │
                      ▼
              SharePoint-Liste TechnikMeldungen  ← Live Pull/Push
```

Kein Genehmigungsfluss, kein Bookings, keine Passwort-Datenbank.

| Schicht | Wert |
|---------|------|
| Lizenz | Modul `betrieb` / Feature `technik-meldungen` |
| Portal-Personas | `lehrer`, `schueler` (Primär), `it` |
| App-Rollen | **Melder** (anlegen, eigene sehen/kommentieren) · **Admin** (Inbox, Status, Löschen, Demo) |

---

## Status-Pipeline

`offen` → `in_arbeit` → `erledigt` (oder `abgelehnt` / wieder öffnen)

Ownership: Melder beim Anlegen; `bearbeitetVon*` + `bearbeitetAm` bei Status/Kommentar.

---

## SharePoint Live

1. **IT:** `tools/sharepoint-liste-technik-meldungen.html` → Site-URL wählen, Liste `TechnikMeldungen` + Spalten anlegen. URL landet in `ms365-technik-meldungen-sp-settings-v1`.
2. **App-Start:** Pull über `resolveSchoolToolSiteUrl('technik')` (`technik-meldungen-sync.js` / `-graph.js`).
3. **Schreiben:** lokal sofort, danach Push Create/Update; Löschen auch in SPO. `spoItemId` wird gemerkt.
4. **403/Fehler:** ruhiger Hinweis, lokaler Stand bleibt (kein Wipe). Demo-Seeds werden nicht nach SPO gepusht.
5. **System (Admin):** „Von SharePoint laden“, Sync-Meta, Link zum Listen-Assistenten.

---

## Shared Quality Bar

- Site: `resolveSchoolToolSiteUrl('technik')` → Settings `ms365-technik-meldungen-sp-settings-v1`
- Stammdaten: Soft-Hydrate, 403 wipe’t lokale Meldungen nicht
- Calm UX: Action-Banner, Quiet Empty, Status-Pipeline; IT-Links nur System/`data-ms365-frontend-it-only`
- Demo local-first: `seedTag` / `localDemoOnly`; Role-Chips nur `?demoRole=1`

---

## Smoke

| Persona | Check |
|---------|--------|
| Schüler | Tool sichtbar; neue Meldung → SPO; nur eigene; kein Status |
| Lehrer | wie Schüler + Zugänge-Hinweise |
| Admin | IT-Inbox, Status → SPO, System/Pull/Demo |
| IT | Listen anlegen → App Pull zeigt schulweite Inbox |
