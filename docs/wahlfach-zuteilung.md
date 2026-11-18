# Wahlfach-Zuteilung (Linz-Auhof)

**Stand:** 2026-11-18  
**Status:** Upgrade auf Shared Quality Bar (Familie `wahl-portal`)  
**Zielgruppe:** Schulleitung / Admin (SLG); Schüler:innen über Anmeldeformular.  
**Familie:** `wahl-portal` (kein approval-planer, kein Bookings, kein Matching-Algorithmus, keine Genehmigungslogik)

---

## Zweck

Schüler:innen melden sich mit Wunschreihenfolge (1.–3. Wahl) für Wahlfach-Angebote einer Runde an. Die Verwaltung prüft Kapazitäten und **teilt manuell zu**. Ergebnislisten und CSV bereiten den Import in die lokale **Access-Schülerakte** vor.

**Kernfluss:** Runde → Angebote (Kapazität) → Schülerwünsche 1.–3. → manuelle Zuteilung → CSV/Access-Export

---

## Abgrenzung

| Tool | Unterschied |
|------|-------------|
| **Schulaktivitäten-Planer** | Genehmigung von Exkursionen; Approvals – andere Familie |
| **Schüler-Stammdaten-Portal** | Einschreibung / Stammdaten |
| **sRDP-Anmeldung** | Matura-Wahlfach mündlich – andere Domäne, gleiche Familie `wahl-portal` |
| **Schulsprecherwahl** | Anonyme Ballots – andere Domäne; SharePoint hier ok (4 Listen) |
| **Wahlfach-Zuteilung** | Kapazität + Wünsche → **manuelle** Zuteilung (kein Auto-Matching) |

---

## Rollenvertrag

| Schicht | Wert |
|---------|------|
| Lizenz | Modul `schulapps` → Tools `wahlfach-zuteilung`, `wahlfach-anmeldung`, `sharepoint-liste-wahlfach` |
| Portal-Persona | Admin-Tool: Lehrkraft/SLG · Schüler-Formular: Schüler · Listen: IT |
| App-Rolle | SLG teilt zu / wertet aus; Schüler trägt Wünsche ein |

Demo-Rollen nur mit `?demoRole=1` (IT-Vorschau Admin-Zugang). Produktiv keine Role-Chips.

---

## Einstieg

| Rolle | URL |
|-------|-----|
| Admin / SLG | [`tools/wahlfach-zuteilung.html`](../tools/wahlfach-zuteilung.html) |
| Schüler-Formular | [`tools/wahlfach-anmeldung.html`](../tools/wahlfach-anmeldung.html) |
| SharePoint-Paket (IT) | [`tools/sharepoint-liste-wahlfach.html`](../tools/sharepoint-liste-wahlfach.html) |

Schüler-Kachel im Portal: **Wahlfach anmelden**. Listen-Setup nur unter System / IT-Chrome, nicht in der Schüler-Primär-UI.

---

## Phasen (sichtbar)

| Phase | Bedeutung |
|-------|-----------|
| Entwurf | Runde vorbereiten |
| Anmeldung offen | Status `offen` + Datum im Fenster |
| Fenster zu / Zuteilung | Geschlossen oder Status `zuteilung` – manuell zuweisen |
| Ergebnis veröffentlicht | Status `veroeffentlicht` – CSV/Access |
| Archiviert | Abgeschlossen |

Admin: Action-Banner „Was jetzt?“ (Anmeldungen prüfen / zuteilen / exportieren). Schüler: Phase-Badge + geschlossenes Formular außerhalb des Fensters.

---

## SharePoint (System of Record)

Site über `resolveSchoolToolSiteUrl('wahlfach')` / Settings `ms365-wahlfach-zuteilung-settings-v1` (Override + Sync in `school-tool-sites`).

Vier Listen (Namen in **System** änderbar):

1. **Wahlfach-Runden** – Schuljahr, Anmeldefenster, Status  
2. **Wahlfach-Angebote** – Bezeichnung, Min/Max, Jahrgänge  
3. **Wahlfach-Anmeldungen** – Schüler, Klasse, E-Mail, Wunsch 1–3  
4. **Wahlfach-Zuteilungen** – endgültiges Ergebnis (+ `AccessExportKey`)

„Paket anlegen“ ist idempotent.

**Live Sync:** Soft-Pull beim Start (Admin + Schüler). Schreiben lokal sofort, dann Soft-Push (`spoItemId` merken): Runde/Angebot/Anmeldung/Zuteilung. Bei 403/Fehler bleibt der lokale Stand (kein Demo-/Cache-Wipe). Demo-Seeds (`wz-demo-auhof-v1` / `*-demo-*`) werden nicht nach SPO gepusht.

---

## Shared Quality Bar (Ist)

| Punkt | Stand |
|-------|--------|
| Site kind `wahlfach` | ja |
| Stammdaten hydrate + local-first | ja |
| Soft 403 | ja |
| SPO Live Sync Pull+Push (4 Listen) | ja (`wahlfach-zuteilung-graph.js` + `-sync.js`) |
| `planner-apps-theme` | ja |
| Calm Action-Banner / Quiet Empty | ja |
| Demo Opt-in `?demo=1` / System-Button | ja |
| Ownership (`schuelerEmail`, `zugeteiltVon`) | ja |
| IT-Setup hinter System | ja |

---

## Ergebnis → Access

Unter **Ergebnis** → **CSV**: Semikolon-CSV (UTF-8 BOM) mit Schuljahr, Angebot, Schüler, Klasse, Wünschen, Zuteilung. Spaltenliste bewusst generisch – Mapping nachziehen, sobald Auhof Screenshots/Feldliste liefert.

---

## Demo local-first

- URL `?demo=1` oder System → „Demo lokal laden“  
- Seed-Tag `wz-demo-auhof-v1`; bei SPO-Fehler bleibt Demo erhalten  
- Kein Auto-Seed mehr nur wegen `?demoRole=1` auf produktivem Stand

---

## Noch mit der Schule klären

1. Wer meldet an (Schul-Account / PIN / Klassenlehrer)?  
2. Überbuchungs- und Nichtzustandegekommen-Regeln  
3. Exakte Access-Importfelder / Screenshots  
4. Eine Runde vs. Nachmeldephase  

---

## Technik

| Thema | Stand |
|-------|--------|
| UI Admin | Calm Shell, Phasen, Action-Banner (`wahlfach-zuteilung-ui.js`) |
| Schüler | Schlankes Formular (`wahlfach-anmeldung-ui.js`) |
| State / Graph / Sync | `wahlfach-zuteilung-state.js`, `-graph.js`, `-sync.js` |
| Daten lokal | `localStorage` `ms365-wahlfach-zuteilung-v1` (Cache; SoR = SPO) |
| Settings | `ms365-wahlfach-zuteilung-settings-v1` + Site-Kind `wahlfach` |
| Zuteilung | manuell + Kapazitätswarnung (Logic DOM-frei) |
| Tests | `tests/wahlfach-zuteilung-*.test.mjs` (inkl. Mapping/Graph) |
