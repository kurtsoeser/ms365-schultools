# Playbooks – Kontrakt & Wellenplan (Saison 1)

## Entscheidungen (verbindlich)

1. **Jährlicher Pfad:** Primär `tools/playbook-schuljahresstart.html`. Organisations-Assistent nur unterstützend.
2. **Ready-Checks:** Hybrid – Checkbox bleibt, Ready-Status `ok|warn|missing|unknown` kommt dazu (`playbook-ready.js`).
3. **Schul-App-Playbooks:** Wave 3 live in `DASHBOARD_PLAYBOOKS` (Kategorie `schulapps`).
4. **Eltern / Elternsprechtag:** zwei getrennte Playbooks behalten.
5. **IT-Nav Schuljahr:** → Schuljahresstart-Playbook (nicht Org-Assistent).

## Deep-Links (`playbook-deeplink.js`)

| Ziel | Muster | Beispiel |
|------|--------|----------|
| Stammdaten-Tabs | `tenant.html#<tab>` | `#lehrer`, `#klassen`, `#faecher` |
| Stammdaten-Setup-Panes | `tenant.html#<pane>` | `#profile`, `#spo-planer` (Planer & Apps) |
| Schularbeiten-Planer | `?view=<id>` | `?view=regeln`, `?view=admin` |
| Freistellung-Setup | `?step=1…7` | `freistellung-setup.html?step=3` |

## Ready-Checks

- API: `evaluateReadyChecks` in `wirePlaybookPage` (`playbook-store.js`)
- DOM: `[data-cp-ready-hint]`, Klassen `pb-step--ready-{status}` (+ Legacy `cp-step--ready-*`)
- Live:
  - Schuljahresstart → `evaluateSchuljahresstartReadyChecks`
  - Blaupause → `evaluateBlaupauseReadyChecks`
  - SA / Freistellung / Apps / Kursteams / Eltern → `playbook-tool-ready.js`

## Wave-Plan

| Wave | Fokus |
|------|--------|
| **0** | Dead-Link, Deep-Link-Kontrakt, Ready-Foundation, App-Stubs |
| **1** | Shared Fluent Playbook-Shell (`playbook-shell.css`) – kein Inline-`cp-*`-CSS |
| **2** | Content Kernpfad (Blaupause, Schuljahr, SA, Freistellung) + Ready-Checks |
| **3** (diese) | App-Playbooks live; Intranet/Kursteams/Eltern gehärtet |

## Shell (Wave 1)

- CSS: `playbook-shell.css` (+ `design-tokens`, `shell-fluent`, `fluent-blocks`)
- Markup: `.pb-page` / `.pb-steps` / `.pb-step` (+ Legacy-Alias `.cp-step` für Hooks)
- Phasen: `.pb-phases` / `.pb-phase__head`; Multi-CTA: `.pb-step__actions`; Pills: `.pb-pill--must|later`
- Persistenz unverändert: `data-cp-*`, `#cpProgress`, `#cpReset`, `wirePlaybookPage`
- App-Playbooks (Wave 3) nutzen dieselbe Shell; Generator: `scripts/_gen-playbook-wave3.mjs`

## Dateien

- Catalog: `dashboard-playbooks-catalog.js` → `DASHBOARD_PLAYBOOKS` (16, inkl. `schulapps`)
- `PLANNED_APP_PLAYBOOKS` leer (deprecated)
- Persistenz: `playbook-store.js`
- Schuljahr-Gates/Ready: `playbook-schuljahresstart-gates.js`
- Blaupause-Ready: `playbook-blaupause-ready.js`
- Tool-/App-Ready: `playbook-tool-ready.js`
- Shell: `playbook-shell.css`
