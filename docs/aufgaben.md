# Aufgaben (interne Zuweisungen)

**Stand:** 2026-11-18  
**Status:** P3 – Shared Quality Bar + SPO Live Sync (SoR)  
**Roadmap:** [`lehrer-portal-roadmap.md`](./lehrer-portal-roadmap.md) § P3  
**Rollenvertrag:** [`rollen-vertrag.md`](./rollen-vertrag.md)  
**SPO Live Sync:** [`spo-live-sync.md`](./spo-live-sync.md) · Tracker [`spo-live-sync-tracker.md`](./spo-live-sync-tracker.md)  
**Lizenz:** Feature `aufgaben` (Paket `betrieb` / `schulapps`)

---

## Familie & Rollen

| Schicht | Wert |
|---------|------|
| Familie | `betrieb` (Tickets: melden → bearbeiten → erledigt) |
| Portal-Persona | Lehrkraft (Alltag); IT nur Setup |
| App-Rollen | **Melder** (anlegen, mir/von_mir) · **Bearbeiter/Admin** (`showAll`, Clear/Demo, System) |

Kein Genehmigungsfluss, kein Bookings, kein Public-Elternformular. Geschwister: Technik-Meldungen, Raumplan, Ausgaben.

---

## Status-Pipeline

`offen` → `in_arbeit` → `erledigt` / `abgelehnt` (Wieder öffnen möglich).

Berechtigt für Status/Kommentar: Zuweisung (E-Mail), Ersteller oder Betrieb-Admin.

---

## Scope-Views

| Scope | Wer sieht was |
|-------|----------------|
| `mir` | Mir zugewiesen |
| `von_mir` | Von mir erstellt |
| `offen` / leer + `showAll` | Nur Admin: schulweit |

---

## SharePoint Live (SoR)

```text
Lehrkraft ── legt Aufgabe an / erledigt zugewiesene
Betrieb-Admin ── schulweite Inbox, Status, System/Pull
Schul-IT ── Liste „Aufgaben“ (Listen-Assistent)
                 │
                 ▼
         SharePoint-Liste Aufgaben  ← Live Pull/Push
```

1. **IT:** `tools/sharepoint-liste-aufgaben.html` → Site + Liste/Spalten; URL in `ms365-aufgaben-sp-settings-v1`.
2. **App-Start:** Pull über `resolveSchoolToolSiteUrl('aufgaben')` (`aufgaben-sync.js` / `aufgaben-graph.js`).
3. **Schreiben:** lokal sofort, danach Push Create/Update; Löschen mit `spoItemId`.
4. **403/Fehler:** ruhiger Hinweis, kein Wipe von Cache/Demo. Demo-Seeds (`seedTag` / `aufgaben-demo-v1`) werden nicht nach SPO gepusht.
5. **System (Admin):** Sync-Meta, „Von SharePoint laden“, Link Listen-Assistent.

### Benachrichtigung
- **Sofort:** mailto bei neuer Zuweisung (Checkbox) und bei erledigt/abgelehnt an Ersteller; Button „Mail“ in der Liste
- **Automatisch:** Power-Automate-Rezept `aufgaben-mail` in [`power-automate-rezepte.html`](../tools/power-automate-rezepte.html) (Trigger auf Liste „Aufgaben“)

### Vorlagen
Karten beim Anlegen: Bitte Unterschrift · Unterlagen nachreichen · Raum prüfen · Elternkontakt (`aufgaben-templates.js`).

### Lehrer-Picker
Zuweisung aus Stammdaten (`setup.teachers`) + optional Entra-Suche; Freitext Name/Mail bleibt möglich.

---

## Shared Quality Bar

- Site: `resolveSchoolToolSiteUrl('aufgaben')` / Settings `ms365-aufgaben-sp-settings-v1`
- Stammdaten: soft Hydrate, 403 wipe’t Cache nicht
- Ownership/Actor: `ErstelltVon*` beim Anlegen; `BearbeitetVon*` bei Status/Kommentar
- Calm UX: Action-Banner, Quiet Empty, Pipeline-Legende; dashed nur System
- Demo: nur `?demo=1` oder System-Button; Seeds mit `aufgaben-demo-v1`; Role-Chips nur `?demoRole=1`
- IT-Setup: `sharepoint-liste-aufgaben.html` – Header `data-ms365-frontend-it-only`, in-App nur System-View

---

## Tool-IDs

- `aufgaben` → `tools/aufgaben.html`
- `sharepoint-liste-aufgaben` → Audience `it`

## Smoke

| Persona | Check |
|---------|--------|
| Melder | Vorlage → Lehrer-Picker → Speichern (+ mailto) → Scope `von_mir` → Item in SPO |
| Admin | „Alle offen“, Status erledigt (+ mailto Ersteller), System Pull, Sync-Meta |
| IT | Listen-Assistent, PA-Rezept Aufgaben Status-Mail |
| Demo | `?demo=1` / `?demoRole=1`; Seeds bleiben lokal |
