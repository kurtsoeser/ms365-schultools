# Schulwahlen (Schulsprecher & SGA Lehrer)

Online-Wahl-Tool im Dashboard als **Schulwahlen** (Tool-ID bleibt `schulsprecherwahl`).  
**Familie:** `wahl-portal` (nicht approval-planer).  
Zwei Wahltypen im selben Tool – Backend und Anonymitätsmodell sind gemeinsam.

## Rollen

| Schicht | Wert |
|---------|------|
| Portal-Personas | `lehrer`, `schueler` |
| App-Rollen | **Admin** (Einrichtung, Phasen, Auszählung, Kundmachung) · **Wahlberechtigte** (Abstimmungsseite) · IT (API/Stammdaten-Setup) |
| Ownership | Meta: `_adminOids` in election-api; UI respektiert `isAdmin` (Lesen ohne Edit/Delete) |
| Ballots | **kein** Author/Ownership – anonym |

## Wahltypen

| Typ (`electionKind`) | Rechtsgrundlage | Punkte | Ergebnis |
|----------------------|-----------------|--------|----------|
| `schulsprecher` | VO Schülervertreter BGBl. Nr. 388/1993 | **6 / 5 / 4** | Schulsprecher:in (absolute Mehrheit der Erstgereihten, sonst Stichwahl) + 2 STV aus Punkten des 1. Gangs |
| `schulsprecher` + `includeSgaDeputies` | dieselbe VO, gleichzeitig SGA-STV | **6 … 1** | wie oben, plus Top 3 nach 3/2/1 = SGA-Schüler-Stellvertreter:innen (ohne bereits gewählte SV-Ämter) |
| `sga_lehrer` | VO Lehrer/Eltern im SGA BGBl. Nr. 389/1993 | **6 … 1** | Top 3 = Lehrervertretung, Platz 4–6 = Stellvertretung; Los nur an Sitzgrenzen; **keine Stichwahl** |

Kandidatenquelle: Schüler-Pool bzw. Lehrer-Pool aus den Stammdaten (local-first + optional Live-Hydrate über Schul-Site).  
Mindestens so viele Kandidaturen wie Reihungen (3 bzw. 6). Die Option **SGA-Schüler-STV** ist nur im Entwurfsstatus einer Schulsprecherwahl umschaltbar.

## Oberflächen

| Seite | Zweck |
|-------|--------|
| `tools/schulsprecherwahl.html` | Admin: Typ wählen, Phasen, Kandidaturen, Zeitfenster, QR, Auszählung, Auswertung |
| `tools/schulsprecherwahl-stimme.html?wahl=<id>` | Abstimmungsformular (QR-Ziel) – schlank, kein IT-Setup |

## Phasen (sichtbar)

Entwurf → offen → ausgezählt → (Stichwahl, nur Schulsprecher) → **Kundmachung**.  
Admin sieht Phasenleiste + Action-Banner **„Was jetzt?“** mit 1-Klick-CTA (eröffnen / auszählen / kundmachen / Link kopieren).

**Kundmachung** ist ein eigener Schritt: Auszählen setzt `publishedResults` noch nicht; erst **Jetzt kundmachen** gibt das Ergebnis frei. Auswertungs-Link: `schulsprecherwahl.html?wahl=<id>&view=results` (ohne Rohstimmen).

**Quorum / Beteiligung:** Hinweis-Banner ab ≥ 50 % der Wahlberechtigten (Schulpraxis). Unterschreitung warnt, sperrt die Auszählung aber nicht.

## Ablauf (Praxis)

1. Admin legt Wahl an (**Schulsprecher:innenwahl** oder **SGA Lehrervertretung**).
2. Kandidaturen aus dem passenden Stammdaten-Pool, Zeitfenster, Anzahl Wahlberechtigter.
3. Wahl eröffnen → QR/Link teilen.
4. Wahlberechtigte melden sich mit **Schul-Konto** an und reihen die Kandidaturen (einmal pro Wahlgang).
5. Admin zählt aus → bei Schulsprecher ggf. Stichwahl → Kundmachung.

## Anonymität & Speicherung – wo liegen die Daten?

**Nein – Stimmen gehören nicht ins Schul-SharePoint.**  
SharePoint setzt bei Graph-Schreiben immer `Author` = die Person. Dann könnten Admins Stimme und Identität verknüpfen.

| Was | Wo | Warum |
|-----|-----|--------|
| Meta (Kandidaturen, Fenster, Status, `electionKind`) | Azure Table `ElectionMeta` (Partition = Schul-Tenant) | alle Geräte dieselbe Wahl |
| „Hat gewählt“ | Azure Table `ElectionVoters` – nur SHA-256-Hash | Einmal-Stimme, kein Stimminhalt |
| Stimmzettel | Azure Table `ElectionBallots` – nur Rankings/Choice | anonym |
| Demo opt-in | `localStorage` (`?demo=1` oder Button) | lokales Testen, Remote unangetastet |
| Personenpools | Stammdaten lokal + optional Live-Hydrate (Site ≠ Stimmen) | Kandidaturen / Begrüßung |

Backend: `backend/election-api`. Schalter: `MS365_ELECTION_API.baseUrl` (`src/shared/election-backend-config.js` / `ms365-config.local.js`).

```js
window.MS365_ELECTION_API = {
  baseUrl: 'https://<function-app>.azurewebsites.net/api/election',
  scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/Election.Access',
  fallbackScope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
  functionKey: '…'
};
```

**Site-URL:** nur für Stammdaten-Hydrate (Default-/Intranet-Site). Niemals „Listen-Site = Stimmen“.

### SPO Live Sync

**⏸ bewusst kein SPO-Listen-SoR** (Tracker #13).  
Pull/Push läuft über `election-api` (Azure Tables), nicht über `*-graph.js`/`sharepoint-liste-*`. Domäne nicht umbiegen: Fake-SPO-CRUD für Ballots wäre ein Privacy-Bruch. SharePoint bleibt auf Personenpools beschränkt.

### Soft Errors

401/403/Netz → ruhige Meldung (`classifyElectionError`), **kein** Wipe von Demo/localStorage oder bereits geladener Liste.

### Deploy (Betreiber)

Resource Group wie License/Kursteams: `rg-ms365schule-kursteams-dev`.

```powershell
cd backend\election-api
.\scripts\ensure-entra-election-scope.ps1
.\scripts\deploy-azure.ps1 `
  -ResourceGroup "rg-ms365schule-kursteams-dev" `
  -FunctionAppName "func-ms365-election-dev" `
  -CreateResources
```

Live (Stand Deploy):
- Health: https://func-ms365-election-dev.azurewebsites.net/api/election/health
- Scope: `api://12e0cfe2-…/Election.Access` (SPA pre-authorized)
- Storage: eigener Account `stms365elec…` (Tables `ElectionMeta` / `ElectionVoters` / `ElectionBallots`)

`ms365-config.js` enthält die `baseUrl` bereits. Neu anmelden, damit Consent für **Election.Access** greift.

Optional später: nur das **fertige Ergebnis** (Ämter) ins Organigramm/`studentCouncil` schreiben – ohne Rohstimmen.

## Demo

- Opt-in: `?demo=1` oder Button **Demo-Auswertung** (Admin)
- Schaltet Client auf lokalen Store (`localDemoOnly`); Remote-API bleibt unangetastet
- Role-Chips nur mit `?demoRole=1` (aktuell keine Demo-Rollen-Chips)

## Code

| Modul | Rolle |
|-------|--------|
| `schulsprecherwahl-kinds.js` | Wahltyp-Templates |
| `schulsprecherwahl-logic.js` | Auszählung (Tests) |
| `schulsprecherwahl-model.js` | Status, Fenster, `electionKind` |
| `schulsprecherwahl-phase.js` | Phasenleiste, nextAdminAction, Demo-Gate |
| `schulsprecherwahl-errors.js` | Soft-Error-Klassifikation |
| `schulsprecherwahl-stammdaten.js` | Pools local-first + Hydrate |
| `schulsprecherwahl-store.js` | lokaler Demo-Store |
| `schulsprecherwahl-service.js` | lokal **oder** election-api (+ Ownership-Flag) |
| `schulsprecherwahl-teachers.js` / `-students.js` | lokale Pool-Reader |
| `backend/election-api/` | Azure Functions + Tables |

## Smoke (kurz)

| Rolle | Check |
|-------|--------|
| Admin | Wahl anlegen → Phasen/Banner → eröffnen → auszählen |
| Schüler:in | Stimme-URL, Login, Reihung, Danke; kein Setup-Panel |
| Lehrer:in (SGA) | SGA-Wahl, 6…1, keine Stichwahl |
| IT | Setup-Panel: API-Hinweis, Stammdaten-Meta; 403 wipe’t nichts |
| Demo | `?demo=1` → lokale Auswertung, Remote unberührt |

## Tests

```bash
npx vitest run tests/schulsprecherwahl-logic.test.mjs tests/schulsprecherwahl-upgrade.test.mjs
cd backend/election-api && npm test
```
