# Schulsprecher:innenwahl

Online-Wahl-Tool nach der österreichischen **Verordnung über die Wahl der Schülervertreter** (BGBl. Nr. 388/1993).

## Wahlverfahren

| Funktion | Wahlpunkte |
|----------|------------|
| Schulsprecher:in | **6** |
| 1. Stellvertretung | **5** |
| 2. Stellvertretung | **4** |

- **Schulsprecher:in:** gewählt, wer auf mehr als der Hälfte der gültigen Stimmzettel mit 6 Punkten (Erstgereihte:r) steht.
- Sonst **Stichwahl** zwischen den beiden häufigsten Erstgereihten (bei Gleichstand: Punkte, sonst Los).
- **Stellvertretungen:** höchste bzw. zweithöchste Punktezahl im **1. Wahlgang**, ohne die Punkte der gewählten Schulsprecher:in.

Optional später: SGA-Stellvertretungen (3 / 2 / 1 Punkte) – im MVP nicht enthalten.

## Oberflächen

| Seite | Zweck |
|-------|--------|
| `tools/schulsprecherwahl.html` | Admin: Kandidaturen, Zeitfenster, QR, Auszählung, Auswertung |
| `tools/schulsprecherwahl-stimme.html?wahl=<id>` | Schüler:innen-Formular (QR-Ziel) |

## Ablauf (Praxis)

1. Admin legt Wahl an, wählt Kandidaturen aus dem **Schüler-Pool** (Stammdaten).
2. Zeitfenster setzen, Anzahl Wahlberechtigter (für Beteiligung).
3. Wahl eröffnen → QR/Link teilen.
4. Schüler:in meldet sich mit **Schul-Konto** an, reiht drei Personen, einmal pro Wahlgang.
5. Admin zählt aus → ggf. Stichwahl → Kundmachung (Auswertung wie Ämter + Wahlgänge).

## Anonymität & Speicherung – wo liegen die Daten?

**Nein – Stimmen gehören nicht ins Schul-SharePoint.**  
SharePoint setzt bei Graph-Schreiben immer `Author` = die Person. Dann könnten Admins Stimme und Identität verknüpfen.

| Was | Wo | Warum |
|-----|-----|--------|
| Meta (Kandidaturen, Fenster, Status) | Azure Table `ElectionMeta` (Partition = Schul-Tenant) | alle Geräte dieselbe Wahl |
| „Hat gewählt“ | Azure Table `ElectionVoters` – nur SHA-256-Hash | Einmal-Stimme, kein Stimminhalt |
| Stimmzettel | Azure Table `ElectionBallots` – nur Rankings/Choice | anonym |
| Demo ohne Backend | `localStorage` | lokales Testen |

Backend: `backend/election-api`. Schalter: `MS365_ELECTION_API.baseUrl` (`src/shared/election-backend-config.js` / `ms365-config.local.js`).

```js
window.MS365_ELECTION_API = {
  baseUrl: 'https://<function-app>.azurewebsites.net/api/election',
  scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/Election.Access',
  fallbackScope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
  functionKey: '…'
};
```

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

## Code

| Modul | Rolle |
|-------|--------|
| `schulsprecherwahl-logic.js` | Auszählung (Tests) |
| `schulsprecherwahl-model.js` | Status, Fenster |
| `schulsprecherwahl-store.js` | lokaler Demo-Store |
| `schulsprecherwahl-service.js` | lokal **oder** election-api |
| `schulsprecherwahl-api.js` | HTTP-Client |
| `backend/election-api/` | Azure Functions + Tables |

## Tests

```bash
npx vitest run tests/schulsprecherwahl-logic.test.mjs
cd backend/election-api && npm test
```
