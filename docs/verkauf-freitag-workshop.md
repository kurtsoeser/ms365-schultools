# Freitag-Workshop – Playbook & Verkaufsbrief

**Stand:** 2026-11-18 (Mi Abend)  
**Ziel:** Alle Schulen: Basis + 3 Kern-Tools (Aktivitäten, Lehrerfreistellung, Matura/sRDP) · Individuell: Wahlfach / Stammdaten · Angebot klar  
**Marke:** MS365-Schultools · Daten im Schul-Tenant  
**Canvas:** `freitag-use-case-prep.canvas.tsx` (Readiness + Playbook + Sales)

### Readiness-Kurzurteil (Deep-Dive 18.11. Abend)

| Tool | Demo | Live / Deploy | Härtester Stolper |
|------|------|---------------|-------------------|
| Schulaktivitäten | grün (~41 Seeds, 46 Tests) | gelb (viel uncommitted) | `?demoRole=1`; Demo-Pfad jetzt `direktion` |
| Lehrerfreistellung | grün (21 Tests) | gelb | `?demoRole=1`; PA nicht live für 6 Schulen |
| Matura / sRDP | grün lokal HAK (27 Tests) | **rot ohne Deploy** | App/Mail untracked → Live-404; Plenum nur HAK |

**Demo-URLs:**  
`schulaktivitaeten-planer.html?demoRole=1` · `lehrer-freistellung-planer.html?demoRole=1` · `srdp-anmeldung.html?demo=1&demoRole=1`

**P0 bis Do:** (1) Release inkl. sRDP **oder** Freitag strikt lokal · (2) Dry-Run aller drei · (3) Bundle/Audience Schulalltag.

---

## 0. Mi–Do Komprimiert (was wirklich zählt)

| Tag | Fokus | Fertig wenn … |
|-----|--------|----------------|
| **Mi** | 3 Kern-Tools 100 % + Setup-Playbook | AKT + LFR + sRDP je unter ~10–12 Min E2E; Setup-Checkliste §4 |
| **Do** | Dry-Run ~100 Min + Individuell | Probe: Setup → AKT → LFR → sRDP → Close; nur Bugfixes |
| **Fr** | Playbook unten | Nie 107 Kacheln; Portal-Ansicht |

**Diese Woche bewusst nicht:** Canvas-Dashboard, BIP, neue Apps, Landing-Redesign, Monolith-Splits.

---

## 1. Preispakete (kundenöffentlich – final für Freitag)

Nicht die 13 Admin-Module vorlesen. **Drei Bündel + Add-ons nach Schmerz.**

| Bundle | Was die Schule bekommt | Preis / Schuljahr | Wann anbieten |
|--------|------------------------|-------------------|---------------|
| **Starter IT** | `core`+`personen`+`schuljahr`+`unterricht` – Stammdaten, Personen, Schuljahr, Kursteams, Hygiene | **390 €** | Nur IT im Raum, noch kein Lehrer-Portal-Bedarf |
| **Schulalltag** | Starter + `intranet`+`schulapps`+`freistellung` – Planer, Freistellungen, Lehrer-Portal | **890 €** | **Hauptangebot Freitag** |
| **Schule Plus** | Schulalltag + `ablauf`+`ausgaben`+`betrieb`+`automationen` – Elternkanal, Ausgaben, Support-Story | **1.690 €** | „Wir wollen alles + ihr richtet es ein“ |

Admin: Schnellwahl unter **Lizenzen → Schildetail** und Referenz unter **Module** (`sales-bundles-catalog.js`).

### Add-ons (nur wenn Schmerz klar)

| Add-on | Wenn die Schule sagt … | Preis / Jahr |
|--------|------------------------|--------------|
| Eltern & Ablauf | WhatsApp/Zettel für Einwilligung, Brief, Krank | **+290 €** (wenn nicht Plus) |
| Ausgaben & Zugang | „Wer hat Schlüssel 12b?“ | **+150 €** |
| Stammdaten-Portal | Einschreibung / Elternformular | **+390 €** |
| **Setup einmalig** | „Richtet ihr das ein?“ | **590–790 €** (je nach Umfang) |

### Psychologische Leitplanken

- Unter ~500 €: oft ohne großen Beschluss → Starter  
- ~800–1.000 €: mit kurzer Begründung → Schulalltag  
- ~1.500–1.800 €: mit Support-Story → Plus  
- Marktanker: Teamsoft ~399 · Vis365 600–900 · Virtualschool ~2.200  

### Abschluss-Satz (wörtlich üben)

> „Wir richten das gemeinsam ein. Wenn ihr damit weiterarbeiten wollt: **Schulalltag 890 € im Jahr** – oder Plus 1.690 € inkl. Elternkanal und Support. Daten bleiben in eurem Microsoft-365-Tenant.“

---

## 2. Freitag-Playbook (Tagesstruktur)

**Kernurteil:** Tenant-Freischaltung ist schnell (Admin vorbereitet). Der Zeitfresser ist **Grunddaten + Verortung** (Intranet-Site / IT-Sicherungsbibliothek). Dafür ~2 Stunden einplanen – danach **drei gemeinsame Tools** für alle.

| Phase | Dauer | Inhalt | Fertig wenn … |
|-------|--------|--------|----------------|
| **0 Warm-up** | 15 Min | Zielbild, Datenhoheit, Preise kurz, Tagesplan | Alle wissen: Setup → 3 Tools → Individuell |
| **1 Basis-Setup** | **~2 Std** | Tenant aktiv · Bundle · Grunddaten · Site/Bibliothek · Listen-Ort | Jede Schule hat Verortung + minimale Stammdaten |
| **2 Tool A** | 40–50 Min | **Schulaktivitäten** gemeinsam einrichten + Testfall | 1 vollständiger Antrag-Flow im Schul-Tenant |
| **3 Tool B** | 40–50 Min | **Lehrerfreistellung** („ich möchte frei haben“) | Lehrkraft kann den Flow selbst noch einmal zeigen |
| **4 Tool C** | 40–50 Min | **Matura / sRDP-Anmeldung** – Profil, Liste Jahr, Testanmeldung | Abschlussjahr-Liste steht; Einstieg klar |
| **5 Individuell** | Rest | Auhof Wahlfach · Linz/Schärding Stammdaten | Extra-MVP oder Follow-up terminiert |
| **6 Close** | 20–30 Min | Angebot, Setup-Preis, Owner, Follow-up Echtdaten | Mündliches Ja oder klarer nächster Termin |

### Die drei Kern-Tools (+ Setup), die „100 % sitzen“ müssen

0. **Basis-Setup** (Tenant schnell, Grunddaten/Site langsam)  
1. **Schulaktivitäten-Antrag** – von allen gewünscht  
2. **Lehrerfreistellung** – Lehrkräfte selbst im Prozess  
3. **Matura / sRDP-Anmeldung** – nach Weihnachten schlagend, jetzt vorbereiten  

### Script Tool A – Schulaktivitäten

1. Liste/Setup auf Ziel-Site  
2. Lehrkraft: Musterantrag anlegen  
3. Direktion: genehmigen  
4. Zweiten Fall: ablehnen oder Nachforderung  
5. Admin: Übersicht / „erledigt“-Kennzeichen  
**Scope-Satz:** „Keine automatische Untis-Eintragung im MVP.“

### Script Tool B – Lehrerfreistellung

1. Liste/Setup + Direktion als Genehmiger  
2. Lehrkraft: Antrag „frei haben“  
3. Direktion: genehmigen/ablehnen  
4. Übersicht / Kalender sichtbar  
5. **Abgrenzung 1 Satz:** Schüler-Freistellung = anderes Tool  
**Wow:** Lehrkräfte erleben den Prozess selbst – nicht nur IT.

### Script Tool C – Matura / sRDP

1. App öffnen (`tools/srdp-anmeldung.html`) – linke Nav wie Aktivitäten  
2. **Demo:** Rollen Schüler → Formular absenden → Zusammenfassung; Lehrer → Meine Prüfungen; Admin → Auswertung  
3. Optional live: Setup-Wizard → Liste `sRDP-Anmeldungen {Jahr}` anlegen → Site in Einrichtung laden  
4. Eine echte Testanmeldung im Schul-Tenant  
5. Kurz klären: Wer darf eintragen (Schüler / Verwaltung)?  
**Verkaufssatz:** „Nach Weihnachten braucht ihr die Anmeldung – Formular, Lehrerblick, Auswertung – und auf Wunsch sofort die Bestätigungsmail aus dem Matura-Postfach.“  
**Scope:** App + Einrichtung + optional Flow-ZIP Bestätigungsmail – keine Klausurplanung, keine zentrale srdp.at-Anbindung.  
**Wow (30 Sek):** `srdp-mail-setup.html` → ZIP → „wie bei euch in der Schule: Mail mit Zusammenfassung aus freigegebenem Postfach“.

### Phase 5 – Individuell (nicht Plenum)

| Schule(n) | Tool | Ziel |
|-----------|------|------|
| BHAK Linz-Auhof | Wahlfach-Zuteilung | Runde + Formular + 3 Anmeldungen (ohne Auto-Matching) |
| HAK.HAS Linz + Schärding | Stammdaten-Portal | Erfassung + geprüft + Export (nur mit Sokrates-Vorlage belastbar) |
| Gmunden / Ried / Rohrbach | — | Bleiben im gemeinsamen Kern (inkl. Matura); Extra nur auf Wunsch |

### Was Freitag **nicht** zeigen / versprechen

- 107 Kacheln / IT-Katalog als Einstieg  
- PA „In Entwicklung“  
- BIP / Bildungsportal  
- Zahlungen, Absenzenbuch, Noten, Intune  
- Automatische Untis-/Sokrates-Sync  
- Automatische Wahlfach-Zuteilung  
- Echtdaten-Produktivsetzung am Freitag (Follow-up)

---

## 3. Demo-Checkliste (Kern zuerst)

Vor Do-Abend einmal **grün** machen:

| # | Tool | Check | Prio |
|---|------|-------|------|
| **0** | Basis-Setup | Tenant + Bundle + Site + Mindest-Grunddaten (+ Abschlussjahr) + Audience | **P0** |
| **1** | Schulaktivitäten | `?demoRole=1` → Demo → Antrag → genehmigen → ablehnen → Untis | **P0** |
| **2** | Lehrerfreistellung | `?demoRole=1` → Antrag → Direktion → Kalender/iCal (PA nur zeigen) | **P0** |
| **3** | Matura / sRDP | `?demo=1&demoRole=1` → Schüler HAK → Lehrer/Admin (kein HTL-App-Formular) | **P0** |
| 4 | Wahlfach (Auhof) | Runde + Schülerformular + 3 Anmeldungen + Admin | P1 Individuell |
| 5 | Stammdaten-Portal | Erfassung → Geprüft → Sokrates-Export | P1 Individuell |

**Bonus nur wenn Zeit:** Elternbrief Magic-Link 60 Sek – nicht Kern.

---

## 4. Einrichtungs-Checkliste („wir richten ein“)

Der Block, der Freitag die ~2 Stunden frisst – vorher ideal schon klären:

1. [ ] Tenant-Lizenz (active/trial) – Admin oft schon vorbereitet → schnell  
2. [ ] Bundle **Schulalltag** (`intranet`+`schulapps`+`freistellung`)  
3. [ ] **Entscheidung:** Welche SharePoint-Site / IT-Sicherungsbibliothek?  
4. [ ] Minimale Grunddaten: Klassen + Lehrkräfte (für Demo/Tools)  
5. [ ] Listen-Setup: Schulaktivitäten + Lehrer-Freistellungen + sRDP-Anmeldungen (Jahr)  
6. [ ] Audience: Portal zeigt Aktivitäten + Lehrerfreistellung + sRDP  
7. [ ] Abschlussjahrgänge in Grunddaten erkennbar (für Matura-Filter)  
8. [ ] Owner Fach + IT-Ansprechpartner  
9. [ ] 30-Min Follow-up für Echtdaten vereinbaren  

**Setup-Paket kommunizieren:** „Einmalig 590–790 € – danach läuft ihr mit dem Jahrespreis.“

**Vor Freitag an IT schicken (1 Frage):** „Auf welcher SharePoint-Site sollen die Listen liegen?“

---

## 5. Verkaufssätze (auswendig)

1. *„Untis plant den Unterricht – MS365-Schultools erledigt die Zettel.“*  
2. *„IT richtet ein – Lehrkräfte erledigen den Alltag im gleichen Portal.“*  
3. *„Daten bleiben in eurem Microsoft-365-Tenant – nicht bei uns.“*  
4. *„Ihr zahlt nur, was ihr nutzt – Starter, Schulalltag oder Plus.“*  
5. *„Drei Prozesse heute live: Schulaktivität, Lehrerfreistellung – und Matura-Anmeldung für nach Weihnachten.“*

---

## 6. Do Dry-Run (Timer 90 Min)

1. **15 Min** Setup-Story (Site, Grunddaten, Bundle)  
2. **25 Min** Schulaktivitäten E2E  
3. **25 Min** Lehrerfreistellung E2E  
4. **25 Min** sRDP/Matura E2E (Profil + Liste + Testanmeldung)  
5. **10 Min** Close mit Preis  
6. Notieren: wo stockt → nur das fixen  

Optional danach: Wahlfach oder Stammdaten – nur wenn die drei Kern-Tools grün.

---

## 7. Mapping Admin-Pakete ↔ Kundenbündel (intern)

| Kundenbündel | Typische `LicenseEntitlements.packages` |
|--------------|------------------------------------------|
| Starter IT | `core`, `personen`, `schuljahr`, `unterricht` |
| Schulalltag | + `intranet`, `schulapps`, `freistellung` |
| Schule Plus | + `ablauf`, `ausgaben`, `betrieb`, `automationen` (+ optional Add-on `stammdaten-portal`) |

### Listenpreise Einzelmodule (Admin / à la carte)

| Modul | Preis / Jahr |
|-------|--------------|
| Basis (Core) | **250 €** |
| Personen | **100 €** |
| Schuljahr | **130 €** |
| Unterricht | **200 €** |
| Kommunikation | **150 €** |
| Intranet & Listen | **130 €** |
| Schul-Apps | **300 €** |
| Freistellungen | **150 €** |
| Automationen | **130 €** |
| Stammdaten-Portal | **390 €** |
| Eltern & Ablauf | **290 €** |
| Ausgaben & Zugang | **150 €** |
| Betrieb & Portal | **130 €** |

Summe Starter-Module ≈ **680 €** (Bundle 390) · Schulalltag ≈ **1.260 €** (Bundle 890) · Plus ≈ **1.960 €** (Bundle 1.690).  
Empfehlung immer Bundle; Einzelmodule nur bei gezielter Nachbuchung.

---

## Verwandte Docs

- Canvas Playbook: `freitag-use-case-prep.canvas.tsx`  
- [`schulaktivitaeten-planer.md`](./schulaktivitaeten-planer.md)  
- [`lehrer-freistellung-planer.md`](./lehrer-freistellung-planer.md)  
- [`wahlfach-zuteilung.md`](./wahlfach-zuteilung.md)  
- [`schueler-stammdaten-portal.md`](./schueler-stammdaten-portal.md)  
- [`srdp-anmeldung.md`](./srdp-anmeldung.md)  
- [`monetarisierung.md`](./monetarisierung.md) · [`license-modules.md`](./license-modules.md) · [`rollen-vertrag.md`](./rollen-vertrag.md)  
