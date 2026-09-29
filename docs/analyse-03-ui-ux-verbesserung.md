# Analyse 03 – UI/UX-Verbesserungspotenzial

**Stand:** 2026-09-27  
**Versprechen:** „Microsoft 365 für die Schule. Einfach. Alles.“  
**Ist:** Mächtige Werkzeugkiste mit ~50 Katalog-Kacheln, gutem Schnellstart, aber uneinheitlicher Tool-Erfahrung.

---

## Kurzfassung

Die größte UX-Chance liegt nicht in neuen Farben, sondern in **Orientierung, Kontext und Vertrauen**:

1. Weniger „Kachel-Wald“, mehr **Playbooks** (Schuljahresstart, Intranet, Eltern, Cleanup).  
2. **Überall gleicher Kontext**: Schuljahr · Domain · Anmeldung · lokal vs. Tenant.  
3. **Einheitliches Bulk-Feedback** (Fortschritt, Fehlerliste, Abbrechen).  
4. Auth/Header/`app.css`-Version und Empty-States auf allen Graph-Tools angleichen.

Bestehende Stärken: Dashboard-Schnellstart, Suche/Favoriten, Cleanup- + Schuljahr-Playbook, Demo-Modus, Onboarding-Welcome.

---

## 1. Ist-Bild der Bedienung

| Ebene | Inhalt | UX-Wirkung |
|-------|--------|------------|
| Dashboard | 5 Schnellstart-Tasks + 7 Kategorie-Tabs + ~50 Kacheln | Gut für Power-User, dicht für Neueinsteiger |
| Einrichtung | Wizard + Experten-Struktur | Zwei mentale Modelle |
| Tools | Je eigene HTML-Seite, teils eigene Auth | Wirkt wie viele Apps |
| Hilfe | Stark (`hilfe.html`) | Nicht überall verlinkt |
| Daten | Browser-lokal + SPO-Backup | Gut erklärt im Footer, im Tool oft unsichtbar |

---

## 2. Priorisierte UX-Probleme

### P1 – Orientierung & Menge

**Problem:** Automationen allein ~9 Kacheln; Intranet mischt Setup-Listen und Planer; Schuljahr-Tab hat nur 1 Kachel.  
**Wirkung:** Nutzer wissen nicht, „wo anfangen“ und was optional ist.  
**Richtung:**

- Katalog-Kacheln hinter **Playbook-Einstiegen** bündeln (Automationen → eine Übersicht + Checkliste).  
- Intranet: „Einrichten“ vs. „Nutzen“ trennen.  
- Schuljahr-Tab füllen: Umbenennen, Merge, Archiv, Datei-Migration als verknüpfte Schritte (nicht alles neue Tools).

### P1 – Kontext fehlt in Tools

**Problem:** `context-bar` / Schuljahr-Chip existiert im System, wird in vielen Tools kaum gezeigt.  
**Wirkung:** Support-Fragen „welches Jahr?“, „nur lokal?“.  
**Richtung:** Persistente Statusleiste in allen Tool-HTMLs:

`Schuljahr 2025/26 · schule.at · angemeldet als … · Daten: lokal (+ SPO-Sync)`

### P1 – Auth / Header-Inkonsistenz

**Problem:**

- `app.css?v=10` (Dashboard) vs. oft `?v=9` (Tools) → Theme/Cache-Drift  
- Manche Tools laden `msal-auth-ui` explizit, andere indirekt  
- Toolbar/Zurück-zum-Dashboard nicht überall gleich  

**Richtung:** Ein gemeinsames Layout-Partial bzw. Build-Inject für Header + CSS-Version + Auth-Widget.

### P2 – Empty States & First Run

**Problem:** Empty-State-Banner nur auf wenigen Seiten (Kursteams, Jahrgangsgruppen, SLG, ARGE-Fachgruppen, …).  
**Richtung:** Standard-Komponente aus `empty-state-ui.js` auf allen Listen-/Graph-Tools: Einrichtung · Demo · Backup · Hilfe-Link.

### P2 – Bulk-Aktionen ohne Standard-Feedback

**Problem:** OneNote, Kursteam-Mitglieder, Sync, Archiv – Fortschritt uneinheitlich.  
**Richtung:** Shared Pattern:

- „N von M“ Progress  
- Abbrechen  
- Fehlerliste + CSV-Export  
- Klartext: „ändert Graph / ändert nur lokal“  
- Nach Lauf: Zusammenfassung Toast + Review-Panel  

### P2 – Wizard vs. Experte unklar

**Problem:** Kursteams = Wizard; Schulstruktur = Expertenmodus hinter `?mode=struktur`; Massenanlage Jahrgang „nur Fußzeile“.  
**Richtung:** Pro Tool oben Modus-Umschalter: „Geführt“ | „Experte“; Experten-Features nie nur in Fußzeile.

### P2 – Mobile / dichte Tabellen

**Problem:** Dashboard bricht gut um; Tool-Zweispaltigkeit und Trees bleiben Desktop-first.  
**Richtung:** Kein Full-Responsive für jeden Tree nötig – aber: Sticky Primary-Action, horizontales Scroll für Tabellen mit Hinweis, Auth-Toolbar kompakter unter 768px.

### P3 – Doppelte / versteckte Einstiege

| Thema | Ist | Besser |
|-------|-----|--------|
| Jahrgang | tool-id `jahrgang` → `jahrgangsgruppen.html` | Ein Name überall |
| ARGE | Archiv vs. Fachgruppen | Nur Fachgruppen im Katalog |
| WTG | kein Katalog, nur Umwege | Link aus Gruppenverwaltung + Playbook |
| Teams-Archiv | Archiv-Redirect | Nur in Cleanup/SS erklären |
| PA-Rezepte | viele Einzelkacheln | Übersicht + Playbook, Einzelseiten sekundär |

### P3 – Hilfe-Anbindung

Hilfe ist gut – Tools sollten immer denselben „?“-Link mit Anker zum Thema haben (`hilfe.html#kursteams` etc.).

---

## 3. Konkrete UI-Verbesserungsvorschläge (ohne Redesign-Zwang)

### 3.1 „Shell light“ (pragmatisch, hoher Gewinn)

Statt SPA-Rewrite:

1. Shared Header-Snippet (Logo, Dashboard, Hilfe, Theme, Auth).  
2. Shared Context-Chip.  
3. Shared Footer-Hinweis lokal/Tenant bei Schreibaktionen.  
4. Einheitliche Toast/Dialog (`app-dialog` / Toast-API schon vorhanden).

### 3.2 Playbook-First auf dem Dashboard

Erweitere Schnellstart um 2–3 **Saison-Playbooks** (Karten, nicht 50 Kacheln):

| Playbook | Schritte (Links auf bestehende Tools) |
|----------|----------------------------------------|
| Schuljahresstart | Einstellungen → Klassen → SLG → Jahrgangsgruppen → Kursteams → Cleanup |
| Intranet aufsetzen | Hub → Lehrerliste → Termine → SA-Listen → PW-Listen |
| Elternkommunikation | Eltern-Verteiler → Bookings → (optional) News/PA |
| Aufräumen | Cleanup-Playbook (existiert) |

Katalog bleibt für Suche/Favoriten, aber sekundär.

### 3.3 Vertrauens-UX bei gefährlichen Aktionen

Besonders SLG-Sync, Merge, Umbenennen, Archiv:

- Default: **Vorschau / Dry-Run** (Diff-Panel schon bei SLG – zum Standard machen)  
- Rote Zusammenfassung: „Es würden X entfernt“  
- Checkbox „Entfernen erlauben“ erst nach Preview  
- Truncation-Warnung blockiert Apply (Analyse 01)

### 3.4 Intranet & Automationen entschlacken

- Tab Automationen: nur „Erst-Setup“, „Übersicht“, „Freistellungen“ + Link „alle Rezepte“.  
- Restliche PA-Kacheln aus Katalog entfernen oder hinter Übersicht.  
- Intranet: Planer (Nutzen) vs. Listen-Setup (Einmalig) visuell trennen.

### 3.5 Microcopy / Sprache

Einheitlich:

- „Schul-Liste“ vs. „Microsoft 365 / Entra“  
- „Anlegen“ vs. „Skript erzeugen“ (Online vs. PowerShell)  
- Kursteams: klarer Hinweis, wann Backend/PS nötig ist  

---

## 4. Was schon gut ist (beibehalten)

- Schnellstart 01–05 + Experten-Details  
- Suche, Tabs, Favoriten/Stern, Drag-Order  
- Demo-Switch + Browser-Backup/SPO im Dashboard-Footer  
- Cleanup-Playbook und Organisations-Assistent als Checklisten-Modell  
- Onboarding-Welcome erklärt Lokalität  

---

## 5. Umsetzungsplan

### Sprint UX-1 – Konsistenz (3–4 Tage)

| Aufgabe | DoD |
|---------|-----|
| CSS-Cache-Bust vereinheitlichen (`app.css` Version) | alle Tool-HTMLs gleiche Version |
| Shared Header + Context-Chip in Top-15-Tools | Schuljahr/Domain sichtbar |
| Empty-State auf allen Graph-Listen-Tools | keine leere Tabelle ohne Hinweis |
| Hilfe-Link mit Anker pro Tool | einheitlich in Toolbar |

### Sprint UX-2 – Vertrauen & Bulk (1 Woche)

| Aufgabe | DoD |
|---------|-----|
| Shared Bulk-Progress-Komponente | OneNote + 1 Sync-Tool nutzen sie |
| Dry-Run-Default bei SLG/JG-Sync Leave | keine Massen-Removes ohne Bestätigung |
| Truncation-Banner Standard | bei `truncated === true` Apply gesperrt |
| Modus-Umschalter geführt/experte (SS + Jahrgangsgruppen) | Expertenfunktionen auffindbar |

### Sprint UX-3 – Information Architecture (1–2 Wochen)

| Aufgabe | DoD |
|---------|-----|
| Playbook-Karten Schuljahresstart + Intranet | Dashboard oben, Katalog bleibt |
| Automationen-Katalog ausdünnen | Übersicht als Hub |
| Schuljahr-Tab mit verknüpften Schritten | Umbenennen/Merge/Archiv verlinkt |
| WTG + Archiv klar aus Gruppenverwaltung | kein „verlorenes“ Tool |
| Benennung Jahrgang vereinheitlichen | ein Label überall |

### Sprint UX-4 – Mobile Light (nach Bedarf)

| Aufgabe | DoD |
|---------|-----|
| Toolbar kompakt &lt;768px | Auth + Primary Action erreichbar |
| Tabellen: horizontal scroll + Hinweis | keine abgeschnittenen Actions |
| Playbooks auf Mobile stapelbar | lesbar ohne Zoom |

---

## 6. Erfolgsmetriken (qualitativ)

- Neue Schule findet „Schuljahresstart“ in &lt; 30 Sekunden.  
- Kein Tool ohne sichtbares Schuljahr/Anmeldestatus.  
- Kein Bulk-Remove ohne Preview.  
- Support-Fragen zu „wo sind meine Daten?“ sinken (Kontext-Chip + Footer).  
- Automationen-Tab wirkt aufgeräumt (≤ 4 Primärkacheln).

---

## 7. Bewusst nicht empfohlen

- Kompletter SPA-Rewrite in einem Rutsch  
- Purple/Glow-Redesign oder generisches Card-Dashboard  
- Alle Tools in iframe-Shell packen (hoher Aufwand, Auth-Komplexität)  
- Weitere Dashboard-Kacheln ohne Playbook-Bündelung  

---

## 8. Abhängigkeit zu anderen Analysen

| Analyse | UX-Bezug |
|---------|----------|
| 01 Fehler | Dry-Run, Truncation-Warnung, Match-UI bei Mehrdeutigkeit |
| 02 Code | Shared Header/Progress braucht Shared-Module; Monolith-Splits erleichtern UI-Extract |
| 04 Ausbau | Neue Tools nur mit Playbook-Einstieg und Context-Chip shippen |

---

## 9. Fortsetzung – Runde 2

Nach dem ersten Nachzug (Playbooks, Automationen, Context-Strip, Shared-Module) liegt der Hebel bei **Adoption und Hierarchie**, nicht bei neuen Kacheln.

→ Konkrete Vorschläge, Prioritäten und DoD: [`analyse-03-runde-2-vorschlaege.md`](analyse-03-runde-2-vorschlaege.md)
