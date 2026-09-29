# Analyse 05 – Bestandsaufnahme & Lückenplan

**Stand:** 2026-09-28 (Nachzug abgeschlossen)  
**Zweck:** Abgleich der vier Analyse-Dokumente mit dem Ist-Code, Priorisierung und Nachweis nach Umsetzung.

---

## 1. Kurzfazit (nach Nachzug)

| Analyse | Vorher | Nachher | Kern |
|---------|-------:|--------:|------|
| **01 Fehler/Bugs** | ~85 % | ~90 % | Truncation-Banner + Dry-Run im Lifecycle verankert; Cap/MSAL-Pilot bewusst später |
| **02 Code-Optimierung** | ~45 % | ~45 % | Kein weiterer Monolith-Split (priorisiert UX/Ausbau); Follow-up Phase D |
| **03 UI/UX** | ~15 % | ~75 % | CSS v=10, Context-Statusleiste, Playbooks, Automationen ≤4, Bulk/Truncation shared |
| **04 Ausbau** | ~10 % | ~80 % | Lifecycle-UI, 3 Playbooks, Sync-Monitor, Vertretung, Räume, Baseline |

---

## 2. Analyse 01 – Status

| Item | Status | Hinweis |
|------|--------|---------|
| K1–K5 | done | unverändert |
| Match-Ambiguity-UI | partial | kein Auto-Match; Banner weiter optional |
| Dry-Run Leave | done* | Lifecycle: Dry-Run Default, Leave-Checkbox |
| Truncation-Banner | done | `truncation-banner.js` + Hilfe `#faq-truncation` |
| Cap konfigurierbar | missing | Follow-up |
| M7/M8 | missing | Follow-up / Phase D |

---

## 3. Analyse 02 – Status

Unverändert bewusst: Bind/tenant/setup/JG bleiben dick. Weitere Splits = Phase D.

---

## 4. Analyse 03 – Status (nach Nachzug)

| Sprint | Status | Evidenz |
|--------|--------|---------|
| UX-1 CSS `app.css?v=10` | done | Tools-HTML |
| UX-1 Context-Statusleiste | done | `context-bar.js` → `#ms365ContextStatus` |
| UX-1 Empty-State | partial→besser | Banner-CSS + bestehende Mounts; Lifecycle neu |
| UX-1 Hilfe-Anker | done | neue Playbook/Lifecycle/Monitor-Artikel |
| UX-2 Bulk-Progress | done | `bulk-progress.js` |
| UX-2 Dry-Run Leave | done | Lifecycle-UI |
| UX-3 Playbook-Karten | done | Dashboard Schnellzone + Katalog |
| UX-3 Automationen ≤4 | done | 4 Primär + Hinweis + versteckte Such-Stubs |
| UX-3 Schuljahr-Tab | done | Playbook, Assistent, Umbenennen, Monitor, Cleanup |
| UX-4 Mobile Light | partial | Context-Status kompakter unter 768px |

---

## 5. Analyse 04 – Status (nach Nachzug)

| Item | Status | Dateien |
|------|--------|---------|
| A1 Schüler-Lifecycle-UI | done | `tools/schueler-lifecycle.html`, `schueler-lifecycle-ui.js` |
| A2 Playbooks Start/Intranet | done | `playbook-schuljahresstart.html`, `playbook-intranet.html` |
| A3 WebUntis-Sync-Monitor | done | `webuntis-sync-monitor.html` + Logic + Tests |
| B1 Räume read-first | done | `raeume-ressourcen.html` |
| B2 Vertretungsplan | done | `sharepoint-liste-vertretung.*` |
| B3 Eltern-Playbook | done | `playbook-eltern.html` |
| B4 Schul-Baseline | done | `schul-baseline.html` |
| C Intune/Sprechtag/SDS | später | – |

Shared: `playbook-store.js`, `bulk-progress.js`, `truncation-banner.js`.

---

## 6. Nachzug – ausgeführt

1. Status-Dokument  
2. UX-1: Context-Statusleiste, CSS, Hilfe  
3. Playbooks + Dashboard-Karten + Automationen/Schuljahr  
4. Lifecycle-UI  
5. WebUntis-Monitor + Tests  
6. Bulk-Progress + Truncation-Banner  
7. Vertretung, Räume, Baseline  
8. `npm test` + `npm run build` (siehe Abschnitt 9)

---

## 7. Qualitätskriterien

- [x] „Schuljahresstart“ in &lt; 30 s auf dem Dashboard (Playbook-Karten + Schnellstart)
- [x] Lifecycle: kein Apply ohne Diff / Dry-Run Default
- [x] Sync-Monitor listet actionable fehlende Besitzer/Matches
- [x] Tool-HTMLs: `app.css?v=10`
- [x] Automationen-Tab: ≤4 Primärkacheln
- [x] Tests grün (772), Build grün

---

## 8. Bewusst später

- SS-`bind.js` &lt;600, tenant-settings Tabs  
- Graph-Apply im Lifecycle (jetzt Deep-Links + Preview)  
- Cap konfigurierbar, MSAL-Pilot Personen  
- Intune / Full-Responsive Trees  

---

## 9. Verifikation

```
npm test
npm run build
```

Neue Tests: `tests/webuntis-sync-monitor-logic.test.mjs`, `tests/ux-shared-helpers.test.mjs`.
