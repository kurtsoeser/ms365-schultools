# Lizenz-Module (Schulen & Features)

## Modell

1. **Tenant-Lizenz** – Status (`active`/`trial`) + Gültig-bis (wie bisher)
2. **Module/Pakete** – welche Dashboard-Tools die Schule nutzen darf
3. **Audience** – wer innerhalb der Schule die Kachel sieht (IT/Lehrkraft/Schüler)

Rollenvertrag (Persona vs. App-Rolle vs. Lizenz, Portal/IT-Backend): [`rollen-vertrag.md`](./rollen-vertrag.md).

SharePoint-Feld an der Schulzeile: `LicenseEntitlements` (JSON).

```json
{
  "v": 1,
  "packages": ["core", "personen", "schuljahr"],
  "toolsInclude": [],
  "toolsExclude": []
}
```

## Defaults

| Fall | Verhalten |
|------|-----------|
| Neue Schule | `core` + `personen` + `schuljahr` |
| Bestehende Schule ohne JSON | Legacy: alle Pakete (kein Bruch) |
| Button „Legacy-Module“ | schreibt explizit alle Pakete für Schulen ohne JSON |

## Zusatzpakete (neu)

| Paket-ID | Label | Inhalt |
|----------|-------|--------|
| `ablauf` | Eltern & Ablauf | `einsammler`, …, `kollegium-umfragen`, …; zusätzlich **`teams-antrag`** (auch in `core`) |
| `ausgaben` | Ausgaben & Zugang | `ausgaben`, `sharepoint-liste-ausgaben` |
| `betrieb` | Betrieb & Portal | `termine-hub`, `aufgaben`, `sharepoint-liste-aufgaben`, `technik-meldungen`, `sharepoint-liste-technik-meldungen`, `raumplan`, `sharepoint-liste-raumplan` (Tools auch in `schulapps`) |

## Runtime

- `GET /license/me` → `entitlements.packages` / `features` / `denied`
- Dashboard blendet Kacheln ohne Feature aus (**IT-Vollkatalog und Schulportal** – gleiche Feature-Liste pro Tenant)
- Tool-Seiten zeigen Overlay „Modul nicht freigeschaltet“
- Kursteams-Backend prüft Feature `kursteams` serverseitig
- Audience (IT/Lehrkraft/Schüler) steuert zusätzlich, **wer** freigeschaltete Tools im Portal sieht; Pakete steuern, **ob** die Schule das Modul hat

## Admin

- Übersicht: Modul-Chips
- Detail: `#licenses/school/<id>` – Umschalter **Pakete** / **Werkzeuge**
  - **Pakete**: ganzes Modul ein/aus (inkl. `requires`-Abhängigkeiten)
  - **Werkzeuge**: gruppiert nach Paket; Paket-Master-Schalter + einzelne Tools via `toolsInclude` / `toolsExclude`
- Katalog-Tab **Module**: Zuordnung Tool → Paket (Master-Katalog)
- Katalog-Code: `feature-catalog` (Backend + `src/shared/license-feature-catalog.js`)
- `priceEuro` / `priceHint` je Paket: Listenpreis Einzelmodul (€/Jahr); Bundles in `sales-bundles-catalog.js` (günstiger als Summe)

### Einzelne Werkzeuge

| Aktion | Wirkung |
|--------|---------|
| Tool an, Paket aus | → `toolsInclude` (Paket-`requires` werden mitgeschaltet) |
| Tool aus, Paket an | → `toolsExclude` |
| Paket an/aus | Overrides werden bereinigt (`pruneToolOverrides`) |

Neue Tools immer im Katalog einem Paket zuordnen und Paket-`requires` setzen (z. B. Schul-Apps → Intranet & Listen).

## SharePoint (Betreiber-Liste)

Liste: `MS365-Schultools-Lizenzen` (Site laut `LICENSE_SITE_WEB_URL`)

| Spalte | Typ | Zweck |
|--------|-----|--------|
| Title | Text | Schulname |
| TenantId | Text | Entra Tenant-ID **oder** `PENDING-<slug>` bis die echte GUID bekannt ist |
| Status / ValidUntil / Domains / Contact / Notes | … | wie bisher; neue Schulen Standard `trial` |
| ItContacts | Mehrzeilig | zusätzliche IT-Kontakte |
| **LicenseEntitlements** | Mehrzeilig | JSON der Module |
| Standort, Strasse, PLZ, Ort | Text | Adresse |
| Direktion, Schuladministration, ItM365Admin, WeitereKontakte | Text/Mehrzeilig | Rollen |
| Telefon, Schulnummer | Text | Kontakt / Kennzahl(en) |
| DomainHinweis, Datenstatus | Text/Mehrzeilig | Meta |
| WorkshopTeilnehmende, WorkshopQuelle, OeffentlicheQuellen, OffenePunkte | Mehrzeilig | Workshop-Quellen |

Katalog-Code: `school-stammdaten-catalog` (Backend + `src/shared/school-stammdaten-catalog.js`).

### Pending-Tenant

- Anlage ohne GUID → automatisch `PENDING-<schulname-slug>`
- Lizenz-Check matched nur echte GUIDs; Pending-Schulen blockieren niemanden fälschlich
- Schul-IT sieht die echte Tenant-ID im **Konto-Menü** (Kopier-Button) und im Lizenz-Overlay („Kein Zugang“)

### Admin-UI (Schulen & Lizenzen)

- Vollbreite Betreiber-Shell
- Schlanke Liste: Schule · Status · Tenant · Gültig bis · Modulanzahl
- Filter: Suche, Status, Tenant + Quick-Chips (Alle / Tenant offen / Trial / Active), Persistenz in localStorage
- Sticky Tabellenkopf, ruhigere Modul-/Bündel-Karten
- Detail-Tabs: **Stammdaten** → Lizenz → Module → Quellen
- Spalten-Panel: Stammdaten gekennzeichnet, Schema aktualisieren
- Neue Schule: Default-Status `trial`

### Workshop-Import

```bash
cd backend/license-api
node scripts/import-workshop-schools.mjs --dry-run
node scripts/import-workshop-schools.mjs --apply
node scripts/import-workshop-schools.mjs --apply --update   # Stammdaten nachziehen
```

Default: Status `trial`, Tenant `PENDING-…`, Gültig bis `2026-12-24`, Excel aus Downloads.

Anlegen / Schema:

1. Admin → **Setup** → „Liste / Spalten aktualisieren“ (Browser, Sites.Manage.All), oder
2. `node backend/license-api/scripts/ensure-entitlements-column.mjs`, oder
3. Nach Deploy: `POST /api/license/admin/ensure-schema` bzw. einmal Schulen laden (legt Entitlements + Stammdaten-Spalten + **`ToolSiteUrl`** an)

**Tool-Site-URL bleibt nach Speichern leer?** Dann fehlt in der Betreiber-Liste (z. B. „Übersicht-Tenants“) die Spalte **`ToolSiteUrl`** (Anzeige: „Tool-Site-URL (Schul-SharePoint)“). Schema wie oben aktualisieren, danach im Schildetail erneut speichern. Ab deploytem Fix legt jedes Speichern die Spalte bei Bedarf selbst an.
