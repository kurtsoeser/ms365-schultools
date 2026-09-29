# Analyse 04 – Fehlende Bausteine & neue Tools

**Stand:** 2026-09-27  
**Leitidee:** Das Versprechen „Microsoft 365 für die Schule. Einfach. Alles.“ soll sich in Tools und Gesamt-App wiederfinden – Fokus Organisation, Verwaltung, Funktionieren.  
**Methode:** Abgleich mit dem bestehenden Dashboard-Katalog (~50 Tools) + Schulalltag (AT/DE).

---

## Kurzfassung

Vieles für **Gruppen, Kursteams, Intranet-Listen, Automationen und Hygiene** ist schon da. Die größten Lücken liegen bei:

1. **Personen-Lifecycle** (Eintritt, Klassenwechsel, Austritt) als geführter Prozess  
2. **Schulbetrieb täglich** (Vertretung, Räume, Sync-Gesundheit Stundenplan)  
3. **Geräte/Intune-Check** (read-only) – heute fast unsichtbar  
4. **Elternkommunikation als Kanal**, nicht nur Verteiler  
5. **Playbooks**, die vorhandene Tools zu „Alles“ verbinden  

Empfehlung: Zuerst Lücken mit **bestehenden Bausteinen** schließen (Playbooks + Lifecycle), dann 2–3 neue Tools mit klarem Graph-Anschluss.

---

## 1. Was bereits abgedeckt ist (nicht neu bauen)

| Bedarf | Vorhanden |
|--------|-----------|
| Klassen / Jahrgang / Merge | jahrgangsgruppen, klassen-merge, klassen-umbenennen |
| Unterrichtsteams | kursteams, templates, einzeln, OneNote |
| Sammelgruppen | SLG, Verwaltung, Klassenvorstände, ARGE-Fachgruppen |
| Personen / Lizenzen / Gäste | personen-verwaltung, lizenzverwaltung, gaeste-* |
| Schuljahr | organisations-assistent |
| Mail / Eltern | postfaecher, verteilerlisten, eltern-verteiler, elternsprechtag-bookings |
| Intranet | Hub, Lehrer-/Termin-/SA-/PW-/sRDP-Listen, Planer |
| Automationen | PA-Rezepte, Freistellung, Seminar, … |
| Hygiene / Cleanup | datenhygiene, leere-gruppen, cleanup-playbook, namenskonvention |
| Struktur / Alle Gruppen | schulstruktur-sync |
| Backup | browser-backup, stammdaten-uebergabe / SPO-Sync |

→ Kein weiteres generisches „Gruppen anlegen“-Tool. Kein zweites Archiv-Tool.

---

## 2. Lücken-Matrix: Schule braucht → Status

| Schulbedarf | Status | Lücke |
|-------------|--------|-------|
| Stammdaten & Struktur | stark | Lifecycle-Prozesse |
| Unterricht / Teams | stark | Sync-Monitor WebUntis |
| Intranet / Formulare | stark | Vertretung / Tagesbetrieb |
| Eltern | mittel | Kanal/News, nicht nur Mail-Verteiler |
| Geräte / Classroom | schwach | Intune/Devices fehlt |
| Räume / Ressourcen | schwach | nur ansatzweise in Postfächern |
| Sicherheit Baseline | mittel | Einzeltools da, kein Gesamt-Check |
| Aufnahme / Sprechtage | teilweise | Schülersprechtag/Aufnahme fehlt |
| SDS / behördliche Exports | schwach | nur Ideen/Doku |
| Playbooks „Einfach“ | teilweise | Cleanup + Schuljahr ja; Rest fragmentiert |

---

## 3. Empfohlene neue Tools / Ausbaustufen

### A – Sofort hoher Nutzen (nächste Tage / 1–2 Sprints)

#### A1. Schüler-Lifecycle-Assistent (neu)

| | |
|---|---|
| **Warum** | Quereinstieg, Klassenwechsel, Abgang – heute verteilt auf Personen, Hygiene, Schuljahr → fehleranfällig |
| **Was** | Geführte Flows: Eintritt · Wechsel · Austritt · Jahrgangswechsel (Hook Org-Assistent) |
| **Daten** | Stammdaten `students`/`teachers`, Graph membership, optional SIS-Import (`school-sis-import`) |
| **UI** | Preview-Diff (join/leave), nie stilles Massen-Remove (Analyse 01) |
| **Aufwand** | mittel (baut auf membership-reconcile + SLG) |

#### A2. Playbook „Schuljahresstart“ & „Intranet“ (Meta, kaum neue Graph-Logik)

| | |
|---|---|
| **Warum** | Macht „Einfach. Alles.“ spürbar ohne 5 neue Monolithen |
| **Was** | Checklisten in `app-data-v2` (wie Cleanup), Fortschritt speichern, Links zu bestehenden Tools |
| **Aufwand** | niedrig–mittel |

#### A3. WebUntis-/Stundenplan-Sync-Monitor (neu, read+report)

| | |
|---|---|
| **Warum** | Import existiert (`webuntis-export-import`); fehlt Dauer-Sicht „letzter Sync, Diff, fehlende Teams“ |
| **Was** | Dashboard: letzter Import, fehlende Kursteams, Lehrer ohne Match, Klassen ohne Team |
| **Anschluss** | Kursteam-Match-Logik + Storage-Timestamps |
| **Aufwand** | mittel |

---

### B – Hoher Nutzen, etwas mehr Scope

#### B1. Räume & Ressourcen (neu)

| | |
|---|---|
| **Warum** | Sekretariat: Räume, Geräte, Buchung – Postfächer-Tool berührt Places nur am Rande |
| **Was** | Übersicht Raum-Postfächer / Room Lists; optional SharePoint-Raumliste; Links Bookings |
| **Graph** | `places`, room lists, ggf. Exchange |
| **Aufwand** | mittel |

#### B2. Vertretungsplan / Kurzinfo (SharePoint-Liste + PA)

| | |
|---|---|
| **Warum** | Täglicher Betrieb; passt zu Intranet-Pattern (Termine, SA) |
| **Was** | Liste anlegen, Ansichten, optional PA-Mail an Lehrergruppe |
| **Aufwand** | niedrig–mittel (Schema + Setup wie andere Listen) |

#### B3. Eltern-Kanal-Playbook (Meta + leichte SPO-Erweiterung)

| | |
|---|---|
| **Warum** | Verteiler + Bookings da; fehlt „wie kommunizieren wir?“ |
| **Was** | Checkliste: Verteiler prüfen → Bookings → Intranet-News/Seiten → PA-Erinnerung |
| **Optional später** | News-Post-Helfer (SharePoint News) |
| **Aufwand** | niedrig zuerst |

#### B4. M365 Schul-Baseline / Secure-Score-Checkliste (read-only)

| | |
|---|---|
| **Warum** | „Wer darf Teams anlegen“, Teilen, Gäste existieren – fehlt Gesamtbild |
| **Was** | Ampel-Checkliste + Deep-Links in bestehende Tools + Secure Score (Graph, falls Scope) |
| **Aufwand** | mittel; klar als „Hinweise“, keine Pseudo-Security-Suite |

---

### C – Sinnvoll mittelfristig

| Idee | Nutzen | Anschluss |
|------|--------|-----------|
| Intune-/Geräte-Schul-Check (read-only) | Endgeräte gehören zu „Alles“ | Intune Graph Compliance |
| Schülersprechtag / Aufnahme-Bookings | Parallel Elternsprechtag | Bookings-Pattern wiederverwenden |
| Klausurraum-Belegung | SA-Planer + Räume | SA-Daten + Places/Listen |
| SDS-CSV-Assistent | Behörden/Insights | Export aus Stammdaten/Struktur |
| Foto/ID für Teams-Gruppen | Klassenbilder | `group-photo-thumb` ausbauen |
| Aktionsprotokoll UX | Vertrauen bei Bulk | `action-log` schon ansatzweise |

---

### D – Bewusst später / Out of Scope

| Thema | Warum nicht jetzt |
|-------|-------------------|
| Voll-LMS / Noten / Zeugnisse | nicht Kern Graph-SPA |
| Zweites Gruppen-Anlege-Tool | Überlappung mit SS/WTG/Kursteams |
| Aggressives Security-Schreiben | Riskant, Consent-Hölle |
| Weitere PA-Einzelseiten ohne Hub | UX-Konsolidierung wichtiger |

---

## 4. Priorisierte Roadmap „nächste Tage“

### Woche 1 – Fundament + Vertrauen

1. Kritische Fixes aus Analyse 01 (Match, Sync-Identität, Truncation, Schuljahr).  
2. Playbook **Schuljahresstart** (nur Checkliste + Links + localStorage-Fortschritt).  
3. UX Context-Chip (Analyse 03) – macht jedes bestehende Tool „runder“.

**Ergebnis:** App wirkt sicherer und einfacher, ohne neues Groß-Tool.

### Woche 2 – Lifecycle

1. **Schüler-Lifecycle-Assistent** MVP: Klassenwechsel + Austritt mit Diff-Preview.  
2. Anbindung Hygiene/Personen (Deep-Links).  
3. Tests für Diff mit Nummern-UPN.

**Ergebnis:** Kernlücke „Verwaltung funktioniert“ geschlossen.

### Woche 3 – Betrieb & Monitor

1. **WebUntis-Sync-Monitor** (Report).  
2. Playbook **Intranet aufsetzen**.  
3. Automationen-Katalog ausdünnen (UX).

**Ergebnis:** Täglicher/periodischer Betrieb sichtbar.

### Woche 4 – Raum & Vertretung (optional parallel)

1. Vertretungsplan-Liste (SPO-Schema).  
2. Räume-Übersicht (read-first).  
3. Eltern-Kanal-Playbook.

---

## 5. Umsetzungsplan je Prioritäts-Tool

### Tool: Schüler-Lifecycle-Assistent

| Phase | Inhalt | DoD |
|-------|--------|-----|
| 0 | Spec: 4 Flows, Felder, Scopes, Dry-Run | Spec 1 Seite |
| 1 | Pure Logic + Vitest (join/leave/class field) | Tests grün |
| 2 | UI Wizard + Diff-Panel | Kein Apply ohne Preview |
| 3 | Graph Apply über Shared Client | Truncation blockiert |
| 4 | Dashboard-Kachel unter Personen + Playbook-Link | auffindbar |

**Geschätzter Aufwand:** 4–6 Entwicklertage.

### Tool: WebUntis-Sync-Monitor

| Phase | Inhalt | DoD |
|-------|--------|-----|
| 1 | Timestamp + Diff-Report aus Storage/Import | Report lokal |
| 2 | Abgleich mit vorhandenen Kursteam-Listen | „fehlende Teams“-Liste |
| 3 | Export CSV + Link Kursteams single-mode | actionable |
| 4 | Dashboard unter Unterricht | sichtbar |

**Aufwand:** 2–4 Tage.

### Meta: Playbooks

| Phase | Inhalt | DoD |
|-------|--------|-----|
| 1 | Datenmodell Checklisten in app-data-v2 | persistiert |
| 2 | UI analog cleanup-playbook | 2 Playbooks live |
| 3 | Dashboard-Hero-Karten | Schnellstart erweitert |

**Aufwand:** 2–3 Tage.

### Tool: Vertretungsplan-Liste

| Phase | Inhalt | DoD |
|-------|--------|-----|
| 1 | Schema + Ensure List (spo-list-kit) | Liste im Intranet |
| 2 | Ansichten (Heute, Klasse) | nutzbar |
| 3 | Optional PA-Rezept | in Übersicht verlinkt |

**Aufwand:** 2–3 Tage.

---

## 6. Beitrag zum Versprechen „Einfach. Alles.“

| Säule | Heute | Nach Roadmap |
|-------|-------|--------------|
| **Einfach** | Viele Kacheln | Playbooks + Context + Dry-Run |
| **Alles (IT/Gruppen)** | stark | + Baseline-Checkliste |
| **Alles (Unterricht)** | stark | + Sync-Monitor |
| **Alles (Personen)** | stark punktuell | + Lifecycle |
| **Alles (Alltag)** | mittel | + Vertretung, Räume |
| **Alles (Geräte)** | fehlend | Intune-Check später |
| **Alles (Eltern)** | Mail/Bookings | + Kanal-Playbook |

---

## 7. Abhängigkeiten / Reihenfolge mit anderen Docs

```
Analyse 01 (Korrektheit) ──┬──► Lifecycle & Sync-Monitor (sonst gefährlich)
Analyse 03 (UX/Playbooks) ─┤
Analyse 02 (Shared Client)─┴──► stabile Graph-Basis für neue Tools
```

**Regel:** Kein neues Schreib-Tool shippen, bevor K2/K3 (Identität + Truncation) gefixt sind.

---

## 8. Entscheidungsvorschläge für dich

Wenn du die nächsten Tage **maximal Wirkung** willst:

1. **Pflicht:** Analyse-01-Fixes + Schuljahresstart-Playbook  
2. **Dann:** Schüler-Lifecycle MVP  
3. **Dann:** WebUntis-Monitor  
4. **Nice-to-have parallel:** Vertretungsplan-Liste  

Wenn du eher **Produktbreite** willst:

1. Playbooks (3 Stück)  
2. Räume + Vertretung  
3. Baseline-Checkliste  
4. Lifecycle  

Beide Wege sind sinnvoll; technisch ist Weg „Korrektheit zuerst“ sicherer.
