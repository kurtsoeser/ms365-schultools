# Schüler-Stammdaten-Portal

Datenschutzkonformes Formular für Erziehungsberechtigte (ohne M365-Login) zur Pflege von Schüler-Stammdaten, inkl. Sekretariats-Admin und Sokrates-Einschreiber-Export.

**Stand Upgrade:** 2026-11-18 · Qualitätsniveau Schulaktivitäten (familienbewusst)  
**Familie:** Hybrid – `ablauf-kommunikation` (Eltern-Formular + Sekretariats-Inbox) + `betrieb` (Status-Pipeline) + `it-setup` (Gateway/Liste/Self-Host)  
**Portal-Audience:** `it` · App-Rollen: Sekretariat/IT (Admin), Eltern (Public/PIN)  
**Kein** approval-planer / Bookings / Wahl-Anonymitäts-Store

## Komponenten

| Rolle | URL |
|---|---|
| Admin (IT/Sekretariat) | `tools/schueler-stammdaten-portal.html` |
| Eltern-Formular | `tools/schueler-stammdaten-formular.html?id=<GUID>&mode=remote` |
| Gateway | `https://func-ms365-ssp-dev.azurewebsites.net/api` (`backend/stammdaten-portal-api/`) |

## Status-Pipeline

```
Vorerfasst → In Bearbeitung → Vollständig → Geprüft → Importiert
```

| Status | Wer | Bedeutung |
|--------|-----|-----------|
| Vorerfasst / In Bearbeitung | Eltern | Entwurf |
| Vollständig | Eltern (Absenden) | Bereit zur Sekretariatsprüfung |
| Geprüft | Sekretariat (Stempel + Ownership) | Freigabe für Sokrates-Export |
| Importiert | System nach XLSX-Export | In Sokrates übernommen |

**Regeln:**

- Eltern können **nicht** mehr speichern, sobald Status `Geprüft` oder `Importiert` ist (UI read-only + API 409).
- Admin-Statuswechsel nur entlang der Pipeline (kein Skip, kein stiller Rücksprung).
- Sokrates-XLSX nur für Status **Geprüft**; danach → Importiert.
- Stempel speichert `ReviewedBy` / `ReviewedByName` / `ReviewedAt`.

## Wo werden Anmeldungen gespeichert?

| Modus | Speicherort | Wann |
|---|---|---|
| **local** (Admin-Setup) | Browser-`localStorage` | Offline-Demo / Pilot am selben PC |
| **remote + `STAMMDATEN_STORE=memory`** | Arbeitsspeicher der Azure Function | Nur Notfall/Debug – Daten weg nach Restart |
| **remote + `STAMMDATEN_STORE=sharepoint`** | SharePoint-Liste **`SchuelerAnmeldungen`** | **Live-Stand** auf `func-ms365-ssp-dev` |

**Live-Pilot (kurtrocks):** Store=`sharepoint`, Site  
`https://kurtrocks.sharepoint.com/sites/MS365-Schultools`, Liste `SchuelerAnmeldungen`.  
Gateway schreibt App-Only; Eltern haben keinen Listen-Link.

Site-URL zusätzlich über Shared `school-tool-sites` Kind **`stammdatenPortal`** (liest `spSiteUrl` aus Portal-Settings).

### Wo lege ich die SharePoint-Liste an?

1. Site im **Schul-Tenant** wählen (Pilot: MS365-Schultools).
2. Admin → **Einrichtung** → **SharePoint-Liste**: Site-URL + Listenname → **Liste anlegen / prüfen**  
   (`POST …/api/ssp-ensure-list`).
3. Berechtigung: nur Sekretariat/Schulleitung (nicht „Jeder“).
4. Azure App Settings: `STAMMDATEN_SP_SITE_URL` + Graph-Credentials + `STAMMDATEN_STORE=sharepoint`.

**Multi-Schule:** Eine Function App = derzeit **eine** Site-URL. Mehrere Schulen → Mapping pro Schul-Slug (nicht „flachbügeln“).

## Betriebsmodell: zentrales Gateway

- Betreiber stellt `func-ms365-ssp-dev` bereit.
- Schule: Config im Admin (+ optional Self-Host-Formular).
- Daten liegen in der SharePoint-Liste der Schule.

## Calm Admin (Shared Quality Bar)

- Action-Banner „Was braucht Aufmerksamkeit?“ (offen / vollständig zu prüfen / bereit zum Export)
- Anmeldungen: Status-Chips + Suche; Quiet Empty statt leerer Roh-Tabelle
- Soft 401/403/Gateway-down: Cache bleibt, verständliche Hinweise (kein Wipe)
- Einrichtung = IT-Chrome (dashed Panels); Alltag = Übersicht/Anmeldungen
- Demo nur **opt-in** (`?demo=1` oder Button) – nie still Remote/Produktiv befüllen

## End-to-End-Pilot (empfohlen)

1. Admin unter Vite öffnen: `http://127.0.0.1:5188/tools/schueler-stammdaten-portal.html`  
   (Proxy `/__ssp-gw`, Dev lädt Admin-Key aus `.admin-key.local.txt`).
2. Sidebar: Chips **remote**, **Gateway**, **Liste**, **sharepoint** sollten grün/aktiv sein.
3. **Neu anlegen** → Titel setzen → anlegen → **PIN notieren**, Link kopieren  
   (Link enthält `mode=remote` und `api=…`).
4. **Formular öffnen** → PIN → Daten speichern/absenden.
5. **Anmeldungen**: Status **Vollständig**, dann Detail → **Als geprüft markieren** (Sekretariats-Stempel).
6. **Sokrates-Export** (XLSX) nur für Status **Geprüft**; danach Status → Importiert.

### Schul-Onboarding (kurz)

1. Site + Graph + `STAMMDATEN_STORE=sharepoint`
2. Liste anlegen (Einrichtung)
3. Schulname/Slug/Preset, Config JSON / Tenant-Datei
4. Optional Self-Host-ZIP
5. Ersten Datensatz wie oben

## Sicherheitsregeln

- **PIN nie in der URL** – nur GUID; PIN getrennt.
- PIN nur als **SHA-256-Hash** (+ Pepper) in SharePoint.
- Fehlversuche → temporäre Sperre; danach Session-Token.
- Admin-Key / Secrets nie ins Eltern-Formular oder öffentliche Tenant-Config.
- Öffentlich: Branding, Feldprofil, `apiBaseUrl`.
- Geprüfte Datensätze: kein stilles Überschreiben durch Eltern-`save`.

## Admin-Defaults (Pilot)

Beim Start setzt das Admin-Tool (falls leer/veraltet):

| Feld | Wert |
|---|---|
| Modus | `remote` |
| API-Base-URL | `https://func-ms365-ssp-dev.azurewebsites.net/api` |
| Site-URL | `https://kurtrocks.sharepoint.com/sites/MS365-Schultools` |
| Listenname | `SchuelerAnmeldungen` |
| Schul-Slug | `ms365schule` |

Admin-Key: lokal per `/__ssp-dev-admin-key` (Vite) bzw. Azure Setting `STAMMDATEN_ADMIN_KEY` / Datei  
`backend/stammdaten-portal-api/.admin-key.local.txt`.

Tenant-Datei: `public/stammdaten-portal-tenants/ms365schule.json` (`mode: remote`).

## Produktion: Gateway + SharePoint

```powershell
cd backend\stammdaten-portal-api
.\scripts\deploy-azure.ps1 `
  -ResourceGroup "rg-ms365schule-kursteams-dev" `
  -FunctionAppName "func-ms365-ssp-dev"
```

**Health:** https://func-ms365-ssp-dev.azurewebsites.net/api/health → `"store":"sharepoint"`

### App Settings

| Setting | Beispiel |
|---|---|
| `STAMMDATEN_STORE` | `sharepoint` |
| `STAMMDATEN_GRAPH_TENANT_ID` | Tenant-GUID |
| `STAMMDATEN_GRAPH_CLIENT_ID` | App-ID |
| `STAMMDATEN_GRAPH_CLIENT_SECRET` | Secret |
| `STAMMDATEN_SP_SITE_URL` | `https://kurtrocks.sharepoint.com/sites/MS365-Schultools` |
| `STAMMDATEN_SP_LIST_NAME` | `SchuelerAnmeldungen` |
| `STAMMDATEN_PIN_PEPPER` | (Deploy) |
| `STAMMDATEN_SESSION_SECRET` | (Deploy, Pflicht in Azure) |
| `STAMMDATEN_ADMIN_KEY` | Admin-Tool (`x-admin-key`) |
| `STAMMDATEN_CORS_ORIGINS` | `*` (+ Host-CORS Vite-Ports / Prod) |

**Pilot Graph:** dieselbe Entra-App wie License-/Kursteams-API (`Sites.ReadWrite.All` bzw. Sites.Selected).

### APIs

| Route | Zweck | Auth |
|---|---|---|
| `POST /unlock` | PIN → Session | id + pin |
| `GET/POST /load` | Datensatz | Bearer |
| `POST /save` | Speichern/Absenden (gesperrt ab Geprüft) | Bearer |
| `POST /ssp-ensure-list` | Liste + Spalten | `x-admin-key` |
| `POST /ssp-enroll` | Anlegen + PIN | `x-admin-key` |
| `POST /ssp-reset-pin` | PIN neu | `x-admin-key` |
| `GET /ssp-list` | Übersicht | `x-admin-key` |
| `POST /ssp-set-status` | Pipeline-Übergang + Stempel | `x-admin-key` |

SharePoint-Spalten: `UniqueGUID`, `AccessPinHash`, `Status`, `PayloadJSON`, `FailedAttempts`, `LockedUntil`, `CreatedAt`, `UpdatedAt`, `ReviewedAt`, `ReviewedBy`, `ReviewedByName`.

## Self-Host

Admin → **Self-Host-ZIP**. Auf Schulwebserver: `index.html?id=<GUID>` + Config mit `mode: "remote"` und `apiBaseUrl` (ohne Admin-Key).

## Feldprofil

Presets: `minimal` | `kern` (Default) | `erweitert`.  
Export: 67 Einschreiber-Spalten (`assets/sokrates/Musterdatei_Einschreiber_V20.xlsx`).

## Demo (local-first, opt-in)

- Button **Demo** / URL `?demo=1` (nur Modus **local**)
- Seed: offen / vollständig / geprüft – PINs in der Browser-Konsole
- Remote: Demo wechselt bewusst auf local (Bestätigung) – keine stillen Produktiv-Füllungen

Backend-Seed (SharePoint, IT):

```powershell
cd backend\stammdaten-portal-api
node scripts\seed-demo-anmeldungen.mjs
```

## Smoke (Abnahme)

| Persona | Check |
|---------|--------|
| IT | Einrichtung: Gateway prüfen, Liste anlegen, Site aus school-tool-sites |
| Sekretariat | Action-Banner → vollständig prüfen → Export nur Geprüft |
| Eltern | PIN-Formular schlank; nach Geprüft nur Lesen |
| Demo | `?demo=1` lokal, Cache bleibt bei Gateway-Fehler |

## Tests

```bash
npm test -- tests/schueler-stammdaten-portal.test.mjs
```

Abgedeckt u. a.: PIN/SVNR/Lockout, Status-Pipeline, Ownership-Stempel, Export-Mapping, Demo-Blueprints.
