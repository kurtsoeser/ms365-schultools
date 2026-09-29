# Analyse 01 – Fehler, Bugs und Datenlücken

**Stand:** 2026-09-27 · Codebasis `MS365schule` · ohne Live-Mandant  
**Ziel:** Risiken finden, die zu **falschen Ergebnissen**, Datenverlust oder problematischen Graph-Aktionen führen können.

---

## Kurzfassung

Die kritischsten Risiken liegen nicht in der UI, sondern in **Matching**, **Mitgliedschafts-Abgleich** und **Schuljahr-Berechnung**. Besonders gefährlich: Substring-Gruppen-Match (`1A` ⊆ `11A`), Sync-Diff ohne `otherMails`/Nummern-UPN sowie stille Truncation großer Gruppen. Dazu kommen Cross-Tool-State-Overwrite und inkonsistente Naming-Normalisierung.

| Priorität | Anzahl | Sofort handeln? |
|-----------|-------:|-----------------|
| Kritisch  | 5      | Ja (Sprint 1)   |
| Mittel    | 10     | Sprint 2–3      |
| Niedrig   | 4      | Backlog         |

---

## 1. Kritische Findings

### K1 – Gruppen-Match: Substring ohne Token-Grenze

| | |
|---|---|
| **Wo** | `src/tools/schulstruktur-sync/schulstruktur-sync-match.js` → `suggestTenantGroupForUnitFromList` |
| **Symptom** | Auto-Vorschlag verknüpft falsche Tenant-Gruppe |
| **Ursache** | Nach Exact-Match greift `includes(uKey)` ohne Mindestlänge/Token-Schutz. Personen-Match hat Schutz (≥3), Gruppen nicht. Beispiel: SOLL `1A` matcht Tenant `Klasse 11A`, weil `"klasse 11a".includes("1a")`. Erster Treffer gewinnt. |
| **Risiko** | Falsche Verknüpfung → Sync/Anlegen an der falschen Klasse; Mitglieder an der falschen Stelle |
| **Tests** | Vorhandene Tests decken Exact/Alias ab, **nicht** den `1A`/`11A`-Fall |

**Fix-Richtung:** Exact → Alias exact → Word-Boundary / Token-Match; Substring nur wenn `uKey.length ≥ 3` **und** nicht als Teil einer längeren Zahl/Klasse (Regex `\b` bzw. Trenner vor/nach). Bei Mehrdeutigkeit: kein Auto-Vorschlag, nur manueller Pick.

---

### K2 – Mitgliedschafts-Sync: Identität nur `mail`/`UPN`

| | |
|---|---|
| **Wo** | `slg-gruppenverwaltung.js` (~679–691), `jahrgangsgruppen.js` (~1470+), `membership-reconcile.js` → `memberEmailFromGraph` |
| **Symptom** | Personen erscheinen als „nur in Graph“ und werden bei Sync entfernt, obwohl sie in der Schul-Liste stehen |
| **Ursache** | Diff nutzt `mail \|\| userPrincipalName`. AT-Education oft: Nummern-UPN `8801@…` + Alias in `otherMails`. `resolveUserByEmail` kennt `otherMails`, der Diff nicht. |
| **Risiko** | Massen-Entfernen aus Sammel-/Klassengruppen |

**Fix-Richtung:** Alle Identitäten in den Diff einbeziehen (`otherMails`, Directory-Index). Truncation (`mem.truncated`) → Sync **abbrechen/warnen**, nie `leave` auf unvollständiger Liste. Regressionstest: Nummern-UPN + Alias in Stammliste.

---

### K3 – Graph-Mitgliederliste still gekürzt (2000 / 40 Seiten)

| | |
|---|---|
| **Wo** | `graph-unified-groups.js` → `fetchGroupMembers`, `fetchAllPagesSimple` |
| **Symptom** | Abgleich wirkt fertig; große Gruppen sind unvollständig |
| **Ursache** | Cap bei 2000 Mitgliedern / 40 Seiten. UI zeigt „gekürzt“ teils an; Sync-Pfade in SLG/JG prüfen `truncated` nicht vor `leave`. |
| **Risiko** | Falsche Joins/Leaves bei großen Sammelgruppen; Directory-Lookup-Cap (~20000) lässt User in großen Tenants aus |

**Fix-Richtung:** `truncated` als Pflicht-Signal; Sync bei Truncation blockieren; Cap konfigurierbar erhöhen oder Count-API vorher prüfen; UI immer mit Warnung.

---

### K4 – Cross-Tool Lost Update auf Struktur-State

| | |
|---|---|
| **Wo** | `schulstruktur-sync-state.js`, `organisations-assistent.js` (`saveStructurePatch`), `app-data-v2.js` (`containerCache`) |
| **Symptom** | Änderungen im Organisations-Assistenten oder zweitem Tab verschwinden nach Speichern im Schulstruktur-Sync |
| **Ursache** | Sync hält `rowsStruktur` im Speicher und schreibt Snapshots; Assistent schreibt parallel nur v2; kein `storage`-Event / Cache-Invalidierung |
| **Risiko** | Verlorene Struktur/Matches vor Schuljahreswechsel |

**Fix-Richtung:** Ein Write-Pfad; vor Save immer frisch laden (optimistic merge oder version field); `storage`-Listener für Multi-Tab; Assistent und Sync denselben State-Helper nutzen.

---

### K5 – Schuljahr-Label ohne September-Grenze

| | |
|---|---|
| **Wo** | `schulstruktur-sync-naming.js`, `app-data-v2.js`, `tenant-settings-ui.js`, `organisations-assistent-logic.js` – jeweils `currentSchoolYearLabel()` |
| **Symptom** | Jan–Aug wird als neues Schuljahr behandelt (März 2026 → `2026/27` statt `2025/26`) |
| **Ursache** | Nur `getFullYear()`, kein Schuljahreswechsel ab 1.9. |
| **Risiko** | Falsches Jahr für Demo, Year-Buckets, Playbook, Rollover |

**Fix-Richtung:** Eine zentrale Funktion in `shared/utils` oder `app-data-v2`: wenn Monat `< 8` (Jan–Aug) → Vorjahr als Startjahr. Alle Duplikate ersetzen. Bestehende Year-Buckets mit Hinweis/Migration.

---

## 2. Mittlere Findings

| ID | Thema | Ort | Risiko |
|----|--------|-----|--------|
| M1 | Inkonsistente Umlaut-/Alias-Normalisierung | Match behält Umlaute; Naming entfernt Nicht-ASCII (`Schüler`→`schler`); Graph-Anlage macht `ä→ae`; Klassen-Umbenennen ohne ae | Doppel-Aliases, Match findet Gruppe nicht |
| M2 | Dual-Storage / Legacy-Divergenz | v1-Keys vs. `app-data-v2`, Quota-Fehler | „Verschwundene“ Matches nach Update |
| M3 | AD-Flags außerhalb Backup | `ms365-*-ad-flags-v1` | Verlust bei Browser-Backup |
| M4 | Personen-Match False Positives | kurze Namen, Score-Gleichstand → Index 0 | Falsche Owner-Zuordnung |
| M5 | Lehrer-Kürzel-Präfix-Match | `kursteam-teacher-match-logic.js` | Falscher Kursteam-Owner |
| M6 | Roster-Filter ohne Counts | `schulstruktur-sync-stats.js` – `ownerCount === -1` ausgeblendet | Leere/gefährliche Gruppen übersehen |
| M7 | Parallele MSAL-Instanzen | klassen-umbenennen, jahrgang-graph, personen-verwaltung | Doppel-Login, Scope-Chaos |
| M8 | Überlappende Klassen-Tools | jahrgang / jahrgangsgruppen / umbenennen / merge / SS / Org-Assistent | Widersprüchliche Alias-Regeln |
| M9 | Klassen-Merge ohne Graph-Schutz | Plan `ok: true` trotz fehlendem Graph-Match | Lokale Klasse ≠ Cloud |
| M10 | Nickname-Konflikt ohne Pagination | `klassen-umbenennen.js` | Späte Rename-Fehler |

---

## 3. Niedrige Findings

| ID | Thema | Hinweis |
|----|--------|---------|
| N1 | Stats-Typ mit Unicode-Bindestrich | ASCII-Legacy unterzählt |
| N2 | `syncEmailsToGroup` N× Graph-Calls | 429-anfällig bei großen Listen |
| N3 | Match-UI: gespeicherte Auswahl vs. Filter | eher UX |
| N4 | Tenant-Cache veraltet | manuelles Neu-Einlesen nötig |

---

## 4. Testlücken (kritisch für falsche Ergebnisse)

| Bereich | Abdeckung | Fehlend |
|---------|-----------|---------|
| Schulstruktur-Sync (match, naming, stats, …) | gut | `1A`/`11A`-Substring; Truncation-Sync |
| Kursteams | gut–mittel | Graph-E2E; Teacher-Prefix-Edgecases |
| Org-Assistent | mittel | Persistenz/Races |
| Jahrgangsgruppen / SLG Sync | schwach | Nummern-UPN + `otherMails` + Truncation |
| Personen-Verwaltung | fehlend | Suche/Lizenz/Create |
| Klassen-Umbenennen / Merge | fehlend | destruktive Pfade |
| `graph-unified-groups` Caps | schwach | Truncation-Flags |

---

## 5. Umsetzungsplan

### Sprint 1 – „Stop the bleeding“ (3–5 Tage)

| Tag | Aufgabe | DoD |
|-----|---------|-----|
| 1 | **K1** Match-Fix + Unit-Tests (`1A`≠`11A`, Token, min. Länge) | Tests grün; Auto-Match keine False Positives in Fixture |
| 1–2 | **K2** + **K3** Diff mit allen Identitäten; Sync bei `truncated` blockieren | Regressionstest Nummern-UPN; SLG/JG zeigen Warnung |
| 2 | **K5** zentrale `currentSchoolYearLabel` (Sep–Aug) | Alle Call-Sites nutzen Shared; Demo zeigt korrektes SJ |
| 3 | **K4** State: Reload-before-save + `storage`-Event | Manueller Multi-Tab-Test bestanden |
| 4–5 | Tests für Merge/Umbenennen-Kernlogik (pure) | Mind. Happy-Path + Konfliktfall |

### Sprint 2 – Stabilität (1 Woche)

1. Naming-Normalisierung vereinheitlichen (eine `normalizeMailNickname`-Funktion).  
2. AD-Flags in `app-data-v2` / Browser-Backup aufnehmen.  
3. Roster-Filter: `-1` = „unbekannt“, nicht „ok“.  
4. Klassen-Merge: bei fehlendem Graph-Match kein stilles `ok`.  
5. MSAL: Tools ohne eigene PCA auf `msal-auth-ui` umstellen (Pilot: personen-verwaltung).

### Sprint 3 – Absicherung (1–2 Wochen)

1. Truncation/Cap-Dokumentation in Hilfe + Tool-Toasts.  
2. Batch-Membership mit besserem Throttling.  
3. Property-Tests für Match (zufällige Klassenkürzel).  
4. Smoke-Checkliste „Schuljahreswechsel“ manuell + teilweise automatisiert.

### Definition of Done (gesamt Phase „Korrektheit“)

- [x] Kein Auto-Match mit Substring-Zahlenfalle  
- [x] Kein Sync-`leave` bei truncated Members oder unaufgelösten Identitäten  
- [x] Ein kanonisches Schuljahr (Sep–Aug)  
- [x] Ein Write-Pfad bzw. Merge-Strategie für Struktur  
- [x] Tests für die fünf kritischen Pfade  

**Umsetzung:** 2026-09-28 (Code + Unit-/Property-Tests + Hilfe-FAQ + `docs/smoke-schuljahreswechsel.md`).

---

## 6. Manuelle Verifikation (Schulen)

1. Tenant mit Klassen `1A` und `11A` → Match-Vorschläge prüfen.  
2. Schüler mit Nummern-UPN + Alias-Mail → SLG-Sync Dry-Run (nur Diff-Panel).  
3. Sammelgruppe > 2000 Mitglieder → Warnung, kein Entfernen.  
4. Org-Assistent ändert Struktur → Sync neu öffnen → Änderung noch da.  
5. Datum simulieren (oder Systemuhr-Hinweis): August vs. September → Year-Label.
