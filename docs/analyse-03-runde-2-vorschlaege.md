# Analyse 03 – Runde 2 (UI/UX-Vorschläge)

**Stand:** 2026-09-28 (umgesetzt)  
**Bezug:** [`analyse-03-ui-ux-verbesserung.md`](analyse-03-ui-ux-verbesserung.md), Ist nach Nachzug in [`analyse-05-bestandsaufnahme.md`](analyse-05-bestandsaufnahme.md)  
**Ziel dieser Runde:** Qualität und Vertrauen vertiefen – nicht noch mehr Kacheln.

---

## 1. Was Runde 1 schon erledigt hat

| Thema | Status |
|-------|--------|
| Playbooks auf dem Dashboard | erledigt (Karten + Katalog) |
| Automationen ≤4 Primärkacheln | erledigt |
| Schuljahr-Tab gefüllt | erledigt |
| Context-Statusleiste (Tools) | erledigt (Dashboard bewusst ausgenommen) |
| Shared Bulk-/Truncation-Module | **gebaut**, kaum **verdrahtet** |
| Hilfe-Anker / CSS-Version Tools | weitgehend erledigt (`v=10`; Dashboard `v=11`) |

→ Runde 2 = **Adoption + Feinschliff**, nicht „noch ein Playbook“.

---

## 2. Symptome, die Nutzer jetzt spüren

1. **Dashboard wirkt voll:** Schnellstart 01–05 + Playbook-Zeile + „Ihr Stand“ + Katalog – alles wichtig, aber ohne klare Hierarchie.  
2. **Vertrauen ist uneven:** Lifecycle hat Dry-Run; SLG / Jahrgangsgruppen / Merge / Archiv noch nicht einheitlich.  
3. **Empty States selten:** nur ~6 Tools mit `empty-state-ui`.  
4. **Bulk-Progress tot im Regal:** Modul da, produktive Tools nutzen es kaum.  
5. **Intranet-Tab mischt weiter** „einmalig anlegen“ und „täglich nutzen“.  
6. **Zwei Logos / Cache-Drift** haben gezeigt: Shell-Konsistenz ist noch fragil.

---

## 3. Vorschläge Runde 2 – Umsetzung

### A – Dashboard-Hierarchie ✅

| # | Status | Nachweis |
|---|--------|----------|
| A2 | ✅ | `index.html` `#dashCatalogFold` standardmäßig zu; Persistenz `ms365-dash-catalog-fold-v1`; öffnet bei Suche |
| A3 | ✅ | Karte 04: Playbook primär, Assistent/Anzeigenamen unter „Weitere Schritte“ |
| A4 | ✅ | Microcopy „Schuljahresstart“ / „Playbook Schuljahresstart“ |

### B – Vertrauens-UX ✅ (Kern)

| # | Status | Nachweis |
|---|--------|----------|
| B1 | ✅ | SLG: Dry-Run Default, Leave-Checkbox, Truncation-Banner + Sync sperren; Shared Bulk-Progress |
| B2 | ✅ | JG: gleiche Trust-Checkboxen (Detail + Sammel), `createBulkProgress`, Abbrechen, Fehlerliste |
| B3 | ⏸ | Merge/Umbenennen/Archiv – bewusst Puffer / Runde 3 |
| B4 | ✅ | Truncation-Guard an SLG + JG Membership-Fetch |

### C – Empty States ✅

| # | Status | Nachweis |
|---|--------|----------|
| C1 | ✅ | ≥20 Graph-Tools mit `data-ms365-empty-state` (u. a. Personen, Gäste, Postfächer, Hygiene, SS, OneNote, Merge, …) |
| C2 | ✅ | Einheitlicher Text „lokal in diesem Browser“ + Demo/Backup/Einrichtung (Shared UI) |

### D – Intranet & Katalog-IA ✅

| # | Status | Nachweis |
|---|--------|----------|
| D1/D2 | ✅ | Intranet-Tab: Unterüberschriften **Einrichten** / **Nutzen** |
| D3 | ✅ | Labels „Klassengruppen“ (Dashboard + Tools) |
| D4 | ✅ | WTG-Link in Schulstruktur-Sync vorhanden |

### E – Shell light ✅

| # | Status | Nachweis |
|---|--------|----------|
| E1 | ✅ | `app.css?v=12` überall |
| E2 | ✅ | Ein Brand-Logo (Regression aus Runde 1) |
| E3 | ✅ | Context-Strip: „angemeldet als …“ / „nicht angemeldet“ |
| E4 | ✅ | Schreibhinweis an Sync-Trust („schreibt nach Microsoft 365“) |

---

## 4. Erfolgscheck Runde 2

- [x] Neue Schule sieht zuerst Schnellstart + Playbooks, Katalog erst nach Klick  
- [x] SLG: kein Leave ohne Preview + Checkbox  
- [x] Mindestens SLG + JG nutzen Shared Bulk-Progress  
- [x] ≥15 Graph-Tools mit Empty-State  
- [x] Intranet-Tab optisch Einrichten ≠ Nutzen  
- [x] Ein `app.css?v=` überall; nur ein Brand-Logo  

---

## 5. Entscheidung (erledigt)

Nutzerwunsch: **Full Runde 2** (A→E).
