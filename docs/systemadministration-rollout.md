# Systemadministration – Rollout auf weitere Tools

**Stand:** 2026-11-19  
**Shared:** `src/shared/system-admin-shell.js` · `src/shared/system-admin-shell.css`  
**Vorbild:** Schulaktivitäten · **bereits umgestellt:** Lehrer-Freistellungen, sRDP-Anmeldung, Wellen A–F  
**Vorlage:** [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md) §5.1

## Muster (kurz)

Nav **Systemadministration** → Tabs **1 Datenquelle · 2 Berechtigungen · 3 Kalender · 4 Wartung · Danger Zone**.  
Vollständiges IT-Setup / Listen-Assistent bleibt eigene Seite; hier nur das, was Direktion/Admin im Alltag braucht.

## Wellen

| Welle | Tools | Familie |
|-------|--------|---------|
| ✅ Done | Schulaktivitäten, Lehrer-Freistellungen, sRDP | approval / wahl |
| ✅ A | Freistellung Schüler | approval-planer |
| ✅ B | Schularbeiten-Planer | register-planer |
| ✅ C | Projektwochen, Elternsprechtag | projekt-buchung |
| ✅ D | Technik-Meldungen, Aufgaben, Raumplan, Ausgaben | betrieb |
| ✅ E | Einsammler, Elternbrief, Krankmeldung, Umfragen | ablauf-kommunikation |
| ✅ F | Schulsprecherwahl, Wahlfach-Zuteilung | wahl-portal |
| — | `sharepoint-liste-*` / PA-Setups | **nicht** umbauen (bleiben `it-setup`) |

## Pro Tool (Checkliste)

1. CSS + Shell-Helper einbinden  
2. `state.systemTab` + Tab-Click  
3. Bestehende Felder mappen (kein Feature-Creep)  
4. IT-Setup-Link unter Datenquelle/Wartung  
5. Smoke: Direktion/Admin sieht Tabs; Lehrer/Schüler nicht  
6. Tests nur wenn Logic neu; Shell-Test `tests/system-admin-shell.test.mjs` deckt Shared ab  

## Agent-Prompt

```text
Propagiere „Systemadministration-Tabs“ von Schulaktivitäten auf <Tool>.
Shared: system-admin-shell.js/.css. Docs: docs/systemadministration-rollout.md + tool-upgrade-vorlage §5.1.
IT-Setup nicht in die Lehrer-Nav ziehen. Domäne nicht umbiegen.
```
