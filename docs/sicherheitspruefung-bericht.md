# Sicherheitsprüfbericht – MS365-Schul-Tools

**Art:** Interne Kontrollprüfung / Bereitschaftsbericht (kein Ersatz für einen beauftragten externen Pentest)  
**Stand:** 23.11.2026  
**Gegenstand:** App `app.ms365.schule`, Landing `ms365.schule`, zugehörige Betreiber-APIs und öffentliches Repo  
**Bezug:** 180-Tage-Plan Initiative I1 (Vertrauensfundament)

---

## 1. Kurzfazit

Die Vertrauenskette der Browser-App ist für den aktuellen Reifegrad **tragfähig**: Abhängigkeiten werden gebündelt, CSP erzwingt `script-src 'self'`, Deploy läuft über FTPS mit SHA-gepinnten Actions, Elternlinks sind HMAC-signiert (Remote-API in Demos aus), Datenschutztexte haben eine Quelle. Der Anmelden-Button fordert nur die Scopes der geöffneten Seite auf der Portal-App. Schreib- und hochprivilegierte Aufrufe laufen über die zweite Entra-App `ms365-schultools-it`. Die Lizenz ist entschieden: Quellcode sichtbar, Produktivnutzung nur mit Schullizenz (`LICENSE`).

**Offen / empfohlen vor breitem Produktivbetrieb:**

1. Externen Sicherheitstest (Web-App + OAuth-Consent + optionale Azure Functions) beauftragen
2. Dev-/Prod-API-Trennung: getrackte Config ohne `-dev`/ohne Key; Deploy injiziert URLs. Azure hat noch nur `*-dev`-Apps – echte Prod-Function-Apps stehen aus
3. Staging-Stufe vor Produktion

---

## 2. Geprüfte Kontrollen

| Kontrolle | Status | Beleg |
|---|---|---|
| Keine fremden Skript-CDNs in App (`src`/`tools`/Root) | OK | Suche leer; `vendor-libs.js` / `msal-loader.js` |
| CSP `script-src 'self'` | OK | `scripts/csp-policy.mjs`, `check-csp` |
| Deploy verschlüsselt + Actions gepinnt | OK | `deploy-ftp.yml`, `deploy-landing-ftp.yml` (FTPS, SHA) |
| Login ohne Admin-Scope-Bundle | OK | `auth-login-scopes.js`, `msal-auth-ui.js` |
| R3 Scope-Bündelung (a4 / c5) | OK | Login persona-schmal auf der Portal-App. Schreib- und hochprivilegierte Aufrufe über `ms365-schultools-it` (`it-app-auth.js`, `itClientId` in `ms365-config.js`). Die Portal-App behält die Schreibscopes in Entra; der Code fordert sie dort nicht an. Zuweisung aus, Zustimmung je Schule einmalig. |
| Lizenz | OK | `LICENSE`: Quellcode sichtbar, Produktivnutzung nur mit Schullizenz |
| Elternlinks HMAC + Fragment | OK | `parent-link-token.js`; Remote-API Demo aus |
| Datenschutz-SSOT | OK | `privacy-copy.js` + Sync |
| Storage-Key-Register | OK | `storage-register.js` + Coverage-Test |
| Trust-Seite öffentlich | OK | `landing/trust.html` |
| AVV-Vorlage | OK | `docs/avv-vorlage.md` |
| Externer Pentest | offen | Auftrag ausstehend |
| Fixture-PII bereinigt | OK (aktueller Tree) | synthetische Fixtures `schule.example`; Historie ggf. noch alt |

---

## 3. Datenfluss (prüfbar)

1. Browser (localStorage / optional IndexedDB)  
2. Microsoft 365 der Schule (Graph / SharePoint / Teams)  
3. Statisches Hosting der App  
4. Betreiber-Minimaldienste (Lizenz, optionale Portale)

Nachweis für Schulen: [Trust-Seite](https://ms365.schule/trust.html) und Netzwerk-Tab auf der App.

---

## 4. Empfehlung externer Test

Scope-Vorschlag für ein externes Audit (1–3 Tage):

- OAuth/MSAL: Consent-Flächen Portal vs. IT, Token-Speicherung, Logout
- CSP und Supply Chain (Build, `vendor/`, Actions)
- SharePoint-Listenrechte / Genehmigungsintegrität (Annahme R5)
- Optionale Azure Functions (Lizenz, Einsammler) inkl. AuthZ und Rate-Limits
- Landing + Trust-Aussagen vs. tatsächlicher Traffic

Ergebnis des externen Tests als Anhang an diesen Bericht hängen und Schulen auf Anfrage bereitstellen.

---

## 5. Freigabe

| | |
|---|---|
| Erstellt | 21.11.2026 · interne Kontrolle zum I1-Stand |
| Fortgeschrieben | 23.11.2026 · R3 (a4 / c5) und Lizenz als umgesetzt geführt |
| Nächste Prüfung | nach externem Pentest oder vor erstem Produktiv-Schuljahr |

---

*Dieser Bericht ist eine ehrliche Bestandsaufnahme. Er ersetzt keine Zertifizierung und keinen unabhängigen Pentest.*
