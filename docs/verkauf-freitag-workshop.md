# Freitag-Workshop – Playbook & Verkaufsbrief

**Stand:** 2026-11-20 (Nacht – Durchführung für den 21.11.)  
**Zeit:** 09:30–16:15  
**Kontext:** **3. Workshop-Tag** der Reihe – kennen sich, Feedback aus den letzten Wochen steckt schon in den Tools  
**Ziel:** Drei Tools **im Tenant eingerichtet** · Rollen-Wow (Direktion / Admin / IT) · Close mit 2-Wochen-Fenster  
**Marke:** MS365-Schultools · Daten im Schul-Tenant · du = Übermittler, nicht „Produktverkäufer“  
**Arbeitsblatt:** [`freitag-use-case-prep.html`](./freitag-use-case-prep.html) (zum Weiterbearbeiten) · Canvas `freitag-use-case-prep.canvas.tsx`

**Entscheidung 20.11.:** Mitnahme 1–3 sind **Schulaktivitäten, Lehrerfreistellung und sRDP-Anmeldung**, alle live. Kalender und Kursteams sind kein Block. Das SharePoint-Formular ist bei jedem Tool eine Minute Kontrast, keine zweite Einrichtung.

---

## Durchführungs-Playbook (Freitag = Tag 3)

### Haltung (wichtigster Rahmen)

Kein Vorstellen, keine Vision-Folien, **keine reine Produktschulung**.

> „Wir haben die letzten Wochen gemeinsam geschärft, was Schulen wirklich brauchen. Heute bauen wir das **in eurem Tenant** – und ihr geht mit mindestens **drei lauffähigen Dingen** raus.  
> Ich bin der Übermittler: Das Prinzip könnt ihr auch ohne meine App nachbauen (SharePoint, Forms, Approvals, Teams). Die App macht denselben Weg **schneller und wiederholbar** – und euer Feedback steckt schon drin.“

| Rolle im Raum | Was sie spüren soll |
|---------------|---------------------|
| **Direktion** | Genehmigen in Minuten statt Zettelchaos; Überblick Kalender/Status |
| **Administration** | Listen, Kalender, ICS – Alltag ohne IT-Tickets |
| **IT / Global Admin** | Consent einmal, Stammdaten/Site, Kursteams & Hygiene als Hebel – weniger Einzelanfragen |

### Versprechen einlösen: 3 mitnehmbare Dinge

| # | Mitnehmen (funktionsfähig) | Typ | Rolle die „Wow“ sagt |
|---|----------------------------|-----|----------------------|
| **1** | **Schulaktivitäten** – ein echter Antrag, Direktion genehmigt, Zeile in der Liste | **Hands-on live** | Direktion + Lehrkraft |
| **2** | **Lehrerfreistellung** – Lehrkraft stellt selbst, Direktion gibt frei | **Hands-on live** | Direktion + Lehrkraft |
| **3** | **sRDP-Anmeldung** – eine Testanmeldung, Lehrerblick, Auswertung | **Hands-on live** (HAK am Beamer) | Admin + Lehrkraft |

**Regel:** Drei Dinge = drei eingerichtete Fälle im Tenant, nicht drei Folien. Wenn die Zeit knappt: #1 vollständig, #2 vollständig, #3 nur eine HAK am Beamer. Kalender, Kursteams, Wahlfach und Stammdaten-Portal wandern in die zwei Wochen danach.

**Formular-Bonus, bei jedem Tool eine Minute:** Liste öffnen, dieselbe Zeile zeigen, Neu-Formular auf- und zuklappen, nichts eintragen. Satz: „Die Zeile liegt bei euch. Das Formular kennt keine Rolle und keine Genehmigung. Deshalb die App.“ Bei sRDP einen Satz mehr, weil „Anmeldung“ nach Formular klingt.

### Dealbreaker vorher klären

| Check | Warum |
|-------|--------|
| **Global Admin / Application Administrator** im Raum (oder erreichbar) | Ohne **Administratorzustimmung** bleibt Login/Consent stecken |
| Pro Schule: IT + Fachperson (Direktion/Admin/Lehrkraft) | IT richtet ein, Fachperson erlebt den Flow |
| URL + WLAN + Demo-Fallback `?demoRole=1` | Login-Hilfe ohne Panik |

### Consent in einfachen Worten (Script ~90 Sek – Tag 3, kein Vortrag)

> „Kurzer Technik-Stopp, dann arbeiten wir: Die App speichert **nichts bei mir**. Sie schreibt in **euren** Tenant. Microsoft fragt einmal: Darf die App im Namen der angemeldeten Person SharePoint-Listen und Gruppen nutzen? Dafür braucht’s oft **Admin-Zustimmung**. Ohne die kommt niemand weiter – das ist Absicht von Microsoft, kein Haken an meinem Tool.“

Alltagssprache nur bei Bedarf: Wer bin ich · Listen lesen/schreiben · Gruppen lesen. Rest später.

**Falls Consent stockt:** Admin consentet → alle neu anmelden. Parallel Demo-Pfad – nicht 20 Min rätseln.

### Strategie-Sätze (wiederholen)

1. *„Ihr habt das mitentwickelt – heute wird’s konkret in eurem Tenant.“*  
2. *„Prinzip geht auch ohne meine App; die App spart euch Wochen Bastelarbeit.“*  
3. *„Die Zeile liegt in eurer SharePoint-Liste. Die Führung – Rolle, Status, Genehmigung – liegt in der App.“*  
4. *„Drei Dinge mitnehmen – danach zwei Wochen echt nutzen, dann entscheiden.“*

### Uhrplan 09:30–16:15 (Tag 3)

| Zeit | Phase | Was du tust | Fertig wenn … |
|------|--------|-------------|----------------|
| **09:30–09:40** | Ankommen & Versprechen | Kein Vorstellen. Brücke zu den letzten Terminen + **3 Mitnahmen** nennen. Rollen kurz ansprechen. | Alle wissen: heute bauen & mitnehmen, Close am Ende |
| **09:40–10:05** | Login & Consent | Jeden reinhelfen. Admin zuerst. Zeit lassen. | Jede Schule ist in der App |
| **10:05–11:00** | Stammdaten gemeinsam | Klassen + Lehrkräfte (WebUntis/Vorlagen). Site klären. **Co-Building-Ton:** wir füllen *eure* Basis. | Minimum für Tool 1 steht |
| **11:00–12:15** | **Mitnahme 1: Schulaktivitäten** | Liste prüfen → ein Antrag → genehmigen → eine Minute SharePoint-Formular → jede Schule wiederholt | **1 echter Antrag liegt** |
| **12:15–13:00** | Pause / Mittag | | |
| **13:00–14:10** | **Mitnahme 2: Lehrerfreistellung** | Genehmiger prüfen → „frei haben“ → Direktion gibt frei → eine Minute Formular | **1 echte Freigabe** |
| **14:10–15:20** | **Mitnahme 3: sRDP-Anmeldung** | Jahr/Liste → eine Testanmeldung → Lehrerblick → Admin → Formular etwas länger | **1 Anmeldung, HAK führt** |
| **15:20–15:35** | Puffer | Hängende Schule nachziehen. Kein neues Tool. | 3er-Liste oder klarer Rest |
| **15:35–16:15** | Feedback & **Close** | Feedback-Würdigung → 3 Mitnahmen laut wiederholen → **2 Wochen Probe** → Bundle nennen → Owner + Termin. Kaufdruck niedrig, Entscheidungspfad klar. | Interesse + Termin + wer entscheidet |

**Puffer-Regel:** Login/Stammdaten dürfen länger dauern. Dann sRDP auf eine HAK am Beamer kürzen. **Close (15:35) ist heilig.** Wahlfach, Stammdaten-Portal und Kursteams nicht mehr im Plenum öffnen.

### Login-Hilfe (Checkliste)

1. URL · 2. Schul-Konto · 3. MFA · 4. Admin-Consent „für Organisation“ · 5. Neu anmelden · 6. >3 Min stocken → Nachbar/Demo, später nachziehen  

### Stammdaten-Minimum

| Muss heute | Nice | Später |
|------------|------|--------|
| Klassen + Lehrer **mit E-Mail** | KV, Abschlussjahr für Matura | Voll-WebUntis, Schüler, SPO-Sync |
| Site / Listen-Ort gewählt | IT-Bibliothek | Perfekte Hygiene |

Keine Exports? 5–8 Klassen + 5 Lehrer manuell – Rest Follow-up. Import: `tools/webuntis-stammdaten-import.html`.

### Scripts (kurz)

**Schulaktivitäten:** Liste prüfen → ein Antrag → Direktion genehmigt → eine Minute SharePoint-Formular → jede Schule wiederholt. Ablehnen, Nachforderung und Kalender nur, wenn der erste Fall schon liegt. Scope: keine Auto-Untis.

**Lehrerfreistellung:** Genehmiger prüfen → „frei haben“ → Direktion gibt frei → eine Minute Formular. 1 Satz: Schüler-Freistellung ist ein anderes Tool. Kalender nur auf Nachfrage, ein Klick.

**sRDP:** Jahr und Liste → eine Testanmeldung → Lehrerblick → Admin → Formular etwas länger stehen lassen. Satz: „Nach Weihnachten braucht ihr das. Das Formular speichert die Zeile, die App führt.“ HAK am Beamer. Keine Klausurplanung, kein srdp.at, Mail-ZIP nur auf Nachfrage.

**Nicht im Plenum:** Kursteams, Wahlfach, Stammdaten-Portal, Termine-Hub. Wenn jemand fragt: „In den zwei Wochen, Owner notieren.“

### Close-Script (Tag 3 – wörtlich üben)

> „Ihr geht mit drei Dingen raus: eine Schulaktivität, eine Lehrerfreistellung, eine sRDP-Anmeldung. Jede Zeile liegt in eurer SharePoint-Liste. Die Führung liegt in der App.  
> Nehmt das zwei Wochen mit, arbeitet mit einem echten Fall, sagt mir was hakt.  
> Dann entscheiden wir ruhig: weiter mit **Schulalltag 890 €/Jahr** – oder schlanker Starter / Plus, je nach Bedarf. Daten bleiben in eurem Tenant. Setup einmalig 590–790 €, wenn ihr wollt, dass wir fertig einrichten.  
> Wer ist Owner bei euch – und welchen Termin nehmen wir in ca. zwei Wochen?“

**Close-Checkliste (Flipchart / Notiz):** Schule · 3 Mitnahmen · Owner Fach · Owner IT · Termin ±2 Wochen · Interesse Bundle (ja/vielleicht/nein)

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

### Abschluss-Satz (wörtlich üben – Tag 3)

> „Drei Dinge mitnehmen, zwei Wochen echt nutzen – dann: **Schulalltag 890 € im Jahr**, oder Starter/Plus je nach Bedarf. Daten bleiben in eurem Tenant. Setup 590–790 € einmalig, wenn ihr wollt, dass wir fertig einrichten.“

---

## 2. Freitag-Playbook (Referenz – Details/Scripts)

**Maßgeblich für morgen:** Abschnitt **Durchführungs-Playbook (Freitag = Tag 3)** oben (Uhrzeiten, 3 Mitnahmen, Haltung, Close).

**Kernurteil:** Tenant-Freischaltung ist schnell. Zeitfresser = **Grunddaten + Verortung**. Danach Hands-on-Mitnahmen, nicht Katalog-Tour.

### Die drei Mitnahmen (+ Setup), die „100 % sitzen“ müssen

0. **Basis-Setup** (Login/Consent, Grunddaten/Site)  
1. **Schulaktivitäten** – Mitnahme 1 für alle  
2. **Lehrerfreistellung *oder* Matura/sRDP** – Mitnahme 2 nach Schulschmerz  
3. **Kalender-Nutzen + Kursteams-Demo** – Mitnahme 3 (Hands-on + IT-Wow) 

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

## 5. Verkaufssätze (auswendig – Tag 3)

1. *„Ihr habt mitentwickelt – ich bündele und übermittle.“*  
2. *„Prinzip geht auch ohne meine App; die App spart Wochen.“*  
3. *„Untis plant Unterricht – wir erledigen die Zettel in eurem Tenant.“*  
4. *„Direktion genehmigt, Admin hat Kalender, IT hat weniger Einzelanfragen.“*  
5. *„Drei Mitnahmen heute – zwei Wochen Probe – dann ruhig entscheiden.“*

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
