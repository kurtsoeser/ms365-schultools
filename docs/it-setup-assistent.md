# IT-Setup-Assistent (einheitliche Schale)

**Stand:** 2026-11-19  
**Familie:** `it-setup`  
**Referenz-Optik:** Schulaktivitäten-Listen (`tools/sharepoint-liste-schulaktivitaeten.html`)  
**Code:** `src/shared/it-setup-shell.js` · `src/shared/it-setup-shell.css`

---

## Ziel

Jedes Schul-Tool braucht irgendwann **Website / Rechte / Paket (oder Flow)**.  
Die **Schale** (Stepper, Status-Chips, Pane-Titel, Footer Zurück/Weiter, Protokoll, Audience-Cards) soll immer gleich aussehen.  
Die **Domäne** (welche Rollen, welche Listen, Flow-ZIP, Kalender, Jahr/Klassen) bleibt tool-spezifisch.

Kurzformel: **Shared Schale + tool-spezifische Panes.**

---

## Pflicht-Elemente (Schale)

| Element | Umsetzung |
|--------|-----------|
| Status-Chips | `tm-perm` Phase + Rechte-Hinweis (`renderItSetupStatusChipsHtml`) |
| Stepper | `tm-glance` + `it-setup-glance--nav` + `it-setup-glance--cols-N` – volle Breite, Designfarbe leicht, **zweizeilig** (Icon + Label; Status-Value nur sr) |
| Pane | `tm-wizard__pane` · `tm-wizard__pane-title` · `tm-wizard__pane-lead` |
| Content-Cards | `tm-panel` (+ optional `tm-panel__head`) |
| Audience | `fr-setup-perm-card` in `fr-setup-perm-grid` (`planner-entra-perms.css`) – volle Breite; `--single` / `--two` / 3-Spalten, **keine** `max-width: 40rem`-Drossel |
| Footer | `tm-wizard__footer` Zurück / Weiter |
| Protokoll | `details.tm-tech` + `div.tm-log` |
| Site-URL | `school-tool-sites.js` |

Wizard-Logik: `createItSetupWizardController({ … })` (Stepper klicken, Validierung vor Sprung, localStorage-Schritt).

---

## Typische Schritt-Muster

### A) Listen-Paket (Referenz AKT / nahe sRDP)

1. **Website** – Site-URL  
2. **Genehmigung** – Admin/Direktion (Gruppe + Personen)  
3. **Einreicher** – Lehrkräfte / Kandidat:innen  
4. **Paket** – Summary + Anlegen + Protokoll  

sRDP ergänzt dazwischen Domänen-Schritte (Klassen, Optionen) – Stepper-Optik bleibt gleich.

### B) Flow-Setup (Lehrer-Freistellung / AKT-PA)

1. Vorbereitung (Onboarding)  
2. Konten / Genehmiger-Mail  
3. Liste  
4. Rechte (Genehmigung + Einreicher + optionale Domäne, z. B. Outlook)  
5. Flow-Paket  

### C) Schlank (Elternbrief)

Kein Mehrschritt nötig: Status-Chips + **Website & Listen** + Paket-Buttons + Protokoll + „Was wird angelegt?“ – gleiche Tokens/CSS.

---

## Neues Tool anbinden (Checkliste)

1. HTML: `tool-modern.css` + **`it-setup-shell.css`** + ggf. `planner-entra-perms.css`  
2. Stepper-Klassen: `it-setup-glance it-setup-glance--nav it-setup-glance--cols-N`  
3. Controller aus `it-setup-shell.js` oder bestehende Tool-Wizard-Logik mit gleichen Labels/Footer  
4. Rechte: bestehende `html*SetupEntraPermGrid` / `createPlannerExtraUsersController` – nicht neu erfinden  
5. Konten-E-Mails: **`wireFlowSetupEmailPickers`** (`src/shared/flow-setup-email-pickers.js`) + „Suchen“-Button – nie nur Freitext ohne Picker  
6. **Flow-Technik-Konto:** nach jedem Listen-Rechte-Schreiben (`apply*PackagePermissions` / `apply*ListPermissions`) Vollzugriff wieder setzen – `flow-service-account-list-access.js` (sonst GetTable/Trigger 403)  
7. Domänen-Panes nur für Tool-Spezial (Kalender, Jahr, Bookings, …)  
8. Smoke: Stepper springt, Weiter validiert, Paket/Flow idempotent, Microsoft-Suche öffnet Dialog  

---

## Bereits angeglichen

| Tool | Seite |
|------|--------|
| Schulaktivitäten-Listen | `tools/sharepoint-liste-schulaktivitaeten.html` |
| sRDP-Liste | `tools/sharepoint-liste-srdp.html` |
| Lehrer-Freistellung Setup | `tools/lehrer-freistellung-setup.html` |
| Schulaktivitäten Flow-Setup | `tools/schulaktivitaeten-setup.html` (Konten-Suche + Rollen wie Listen-Tool) |
| Elternbrief-Listen | `tools/sharepoint-liste-elternbrief.html` (schlank) |

## Nächste Kandidaten

Projektwochen, Schularbeiten-Listen, Schulsprecherwahl, Freistellung-Schüler-Setup – gleiche CSS/Controller + E-Mail-Suche, Domäne belassen.
