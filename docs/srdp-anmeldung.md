# sRDP-Anmeldung – Umsetzungsplan

**Stand:** 2026-11-18  
**Status:** Familie `wahl-portal` – Shared Quality Bar + Calm UX; **Anmeldephase offen/geschlossen**, Doppelanmeldung gesperrt, Frist in Systemadmin editierbar  
**Zielgruppe MVP:** HAK (+ Profile HTL/HLW/BAFEB/AHS im Setup)  
**UI:** Vanilla JS/ESM – **Online-App** mit linker Nav + Guided Wizard  
**Freitag-Workshop:** Tool C – Matura/sRDP live (Demo lokal / Setup + Testanmeldung + Rollenansichten)  
**Demo laden:** Systemadministration → „Demo lokal“ (oder `?demo=1`); Role-Chips nur mit `?demoRole=1`

Verwandt: [`schulaktivitaeten-planer.md`](./schulaktivitaeten-planer.md), Stammdaten, SharePoint-Listen.

Quellen Varianten-Logik: [hak.cc – RDP](https://www.hak.cc/pruefungen-abschluss/rdp), BMB sRDP BHS; srdp.at = zentrale Klausuren, nicht Variantenmatrix.

---

## 1. Zielbild (Phase 4)

```text
Schüler / Lehrer / Admin ──MSAL──► sRDP-Anmeldung App (linke Nav)
                                      │ Graph Sites.ReadWrite.All
                                      ▼
                             SharePoint: sRDP-Anmeldungen {Jahr}
IT / Admin ──────────────────────────► Einrichtung (bisheriger Wizard)
```

| Rolle | Was sie sieht |
|-------|----------------|
| **Schüler** | Sauberes Anmeldeformular → sofortige **Zusammenfassung** der Wahl |
| **Lehrer** | **Meine Prüfungen** (wo ich DA/LFS/BKO/Wahlfach bin) + ggf. **Meine Klasse** (KV) |
| **Admin / Sekretariat** | Gesamtübersicht, Auswertung nach Klasse/Variante/Wahlfach |
| **IT** | Einrichtung: Site, Jahr, Klassen, LFS/Wahlfach/Seminar, Rechte, Liste anlegen |

**System of Record:** SharePoint-Liste pro Terminjahr (unverändert). Die App ist die bevorzugte Oberfläche; das SPO-Listenformular bleibt Fallback.  
**SPO Live Sync:** Pull beim Start/Reload · Create pusht Anmeldung · Admin-Nacharbeit (`NacharbeitErledigt`) pusht per PATCH · Demo-Seeds/`localDemoOnly` nicht nach SPO · 403 wipe’t Cache/Demo nicht · Sync-Meta in Systemadministration.

---

## 2. Entscheidungen (fest)

| Thema | Entscheidung |
|-------|--------------|
| Formular | **SharePoint-Listenformular** (modern), kein Microsoft Forms |
| Stammdaten | **tenant-settings**: Klassen (nur Abschlussjahrgänge), Lehrer; Fächer-Codes wo sinnvoll (LFS) |
| Wahlfach / Seminar | Setup-Wizard: manuell pflegen; Seminar = Choice **mit Freitext** |
| Varianten-Logik | Pflicht – HAK Variante 1–3; Formular/Validierung abhängig davon |
| Listenlebenszyklus | **Jedes Jahr neu** (`sRDP-Anmeldungen 2026`, …) |
| Item-Titel | automatisch `Nachname, Vorname` |
| Rechte | Broken Inheritance + Gruppen **beim Setup setzen** |
| Schulformen | **HAK, HTL, HLW, BAFEB, AHS** – Profilwahl im Wizard |
| Backend | Kein eigenes Azure für CRUD; Graph wie bei anderen SPO-Listen |
| UI in den Tools | Setup + Health-Check; Anmeldung selbst läuft in SharePoint |

---

## 3. HAK – Prüfungsvarianten (Referenz)

Sieben Prüfungsteile: Diplomarbeit + (3 Klausuren + 3 mündlich) **oder** (4 Klausuren + 2 mündlich).

| | Variante 1 | Variante 2 | Variante 3 |
|---|------------|------------|------------|
| **Diplomarbeit** | ja | ja | ja |
| **schriftlich** | D (5h), BFK (6h), **LFS (5h)** | D (5h), BFK (6h), **AM (4,5h)** | D, BFK, **LFS**, **AM** |
| **mündlich** | BKO, **AM**, Wahlfach | BKO, **LFS**, Wahlfach | BKO, Wahlfach |

Implikation für die Anmeldung (App-Wizard, profilgesteuert):

- Variante zuerst als Karten (schriftlich/mündlich) → dann Prüfungsgebiete.
- **Schriftlich** (zentral/schulisch): kein Prüfer.
- **Mündlich**: Fach wo nötig + **Mehrfach-Lehrer** (`Name1; Name2`), mind. 1.
- **LFS**-Code in allen HAK-Varianten; **Lehrer LFS nur**, wenn LFS mündlich (Var. 2).
- **Lehrer AM** nur bei mündlichem AM (Var. 1); Spalte `LehrerAM`.
- **BKO** / **Wahlfach** immer mündlich → Multi-Lehrer.
- Seminar-Pflicht unverändert, wenn Wahlfach mit „Seminar“ beginnt.
- Gleiche Engine für **HTL / HLW / BAFEB / AHS** (Profile + Part→Feld-Mapping); AHS nutzt ABA/VWA-Felder.

Offizielle Wahlfach-Bezeichnungen (Wizard-Defaults, Schule kann kürzen/erweitern): Religion/Ethik, Kultur, Geschichte und …, Geografie und …, Naturwissenschaften, Recht, Volkswirtschaft, Berufsbezogene Kommunikation in der LFS, Mehrsprachigkeit, Wirtschaftsinformatik, Seminar …, Freigegenstand …, (Sonderfälle Slowenisch/Deutsch nur zweisprachige HAK).

---

### 3.1 ABA-Portal → Abschlussarbeiten

IT/Admin lädt den Excel-Export aus dem ABA-Portal hoch (SheetJS) – im **IT-Setup** (`sharepoint-liste-srdp.html`, Schritt 4) und in der App-Systemadministration. Spalten: Titel, Schüler/in, Klasse, Betreuer/in, Status. Team-Zellen werden **pro Person** aufgelöst und in `sRDP-Abschlussarbeiten {CODE} {Jahr}` geschrieben (Replace pro Terminjahr). „Listen anlegen“ erzeugt Anmeldungs- **und** Abschlussarbeiten-Liste; „Listen prüfen“ prüft beide. Musterdatei: `assets/srdp/ABA-Portal-Demo-Diplomarbeiten.xlsx` (Demo-Schüler:innen inkl. Lena Bauer, Betreuer Mag. Müller &amp; Co.). Beim Anmelden Match auf Vorname/Nachname (+ Klasse) → Prefill Titel + Betreuer (BHS: Diplomarbeit; AHS: ABA/VWA).

---

## 4. Datenquellen

### 4.1 Aus tenant-settings (bereits vorhanden)

| Entität | Nutzung |
|---------|---------|
| `classes[]` | Choice **Klasse** – Wizard zeigt **nur Abschlussjahrgänge** (z. B. `year === 5` bzw. Schulform-Regel) |
| `teachers[]` | Choices: Betreuung DA, LFS, BKO, Wahlfach (Anzeigename) |
| `subjects[]` | Vorschlagsliste **LFS**-Kürzel (ENWS, FRWS, …) – Wizard kann übernehmen/streichen |

Keine parallelen Lookup-Listen „Klassen“ / „Lehrer“ auf SharePoint.

### 4.2 Nur im Wizard (pro Termin)

| Einstellung | Beispiel |
|-------------|----------|
| Schulform-Profil | `HAK` (später `HTL`, …) |
| Terminjahr / Listenname | `2026` → Liste `sRDP-Anmeldungen 2026` |
| Site-URL | Intranet aus Setup |
| Klassen-Filter | welche Klassen dürfen wählen |
| LFS-Optionen | aus Fächern + manuell |
| Wahlfach-Optionen | Defaults + manuell |
| Seminar-Optionen | manuell (oder Freitext-Spalte) |
| Optional: Farben | Column Formatting für Klasse / Variante / Wahlfach |

Persistenz der Wizard-Defaults lokal (Browser / Backup-Key analog anderer Tools), damit nächstes Jahr nur Jahr + Klassen angepasst werden.

---

## 5. SharePoint-Liste (MVP HAK)

**Titel:** `sRDP-Anmeldungen {CODE} {Jahr}` bzw. AHS `sRP-Anmeldungen AHS {Jahr}`  
(HAK-Legacy ohne Code wird beim Suchen noch erkannt.)  
**Template:** `genericList`

### 5.1 Spalten

| Interner Name | Anzeige | Typ | Pflicht | Hinweis |
|---------------|---------|-----|---------|---------|
| `Title` | Titel | Text | ja | **automatisch** `Nachname, Vorname` |
| `Klasse` | Klasse | Choice | ja | nur **Abschlussjahrgänge** aus Stammdaten |
| `Nachname` | Nachname | Text | ja | |
| `Vorname` | Vorname | Text | ja | |
| `TitelDiplomarbeit` | Titel Diplomarbeit | Text (mehrzeilig ok) | ja | |
| `BetreuungslehrerDA` | BetreuungslehrerIn Diplomarbeit | Choice | ja | Lehrer |
| `Variante` | Variante | Choice | ja | `Variante 1` / `2` / `3` |
| `LFS` | LFS | Choice | ja | bei HAK immer (siehe §3) |
| `LehrerLFS` | LehrerIn LFS | Choice (`allowTextEntry`) | bedingt | nur wenn LFS mündlich; Multi via `Name; Name` |
| `LehrerBKO` | LehrerIn BKO mündlich | Choice (`allowTextEntry`) | ja | Multi |
| `LehrerAM` | LehrerIn AM mündlich | Choice (`allowTextEntry`) | bedingt | nur HAK Var. 1 (AM mündlich); Multi |
| `WahlfachMuendlich` | Wahlfach mündlich | Choice | ja | |
| `Seminar` | Seminar | Choice **mit Freitext** (`allowTextEntry: true`) | bedingt | Pflicht wenn Wahlfach Seminar; Optionen + eigene Eingabe |
| `LehrerWahlfach` | LehrerIn Wahlfach | Choice (`allowTextEntry`) | ja | Multi |
| `Bestaetigung` | Bestätigung Anmeldung | Boolean / Choice Ja | ja | |
| `Schulform` | Schulform | Choice (hidden) | – | Default `HAK` für spätere Profile |
| `TerminJahr` | Terminjahr | Zahl/Text | – | z. B. 2026 |
| optional `PruefplanKurz` | Prüfplan (Info) | Text berechnet/Beschreibung | – | z. B. „schriftlich: D, BFK, LFS · mündlich: BKO, AM, Wahlfach“ |

Referenzen als **Choice-Text** (Namen/Codes), keine Person-Felder im MVP – entspricht den Screenshots mit Pills.

### 5.2 Formular-Logik (Listenformular)

SharePoint Modern Form – per Setup gesetzt:

1. **ClientFormCustomFormatter:** Header mit Live-Name (`Nachname, Vorname`) und Prüfplan zur Variante; Body in Sektionen; Footer-Hinweis.  
2. **Seminar:** möglichst `ConditionalShowFormula` (sichtbar wenn `WahlfachMuendlich` mit „Seminar“ beginnt). SPO typed Field-MERGE lehnt die Eigenschaft oft ab – Setup setzt sie best-effort; native Liste kann Seminar dann dauerhaft zeigen (App-Wizard blendet korrekt aus).  
3. **Title / Schulform / Terminjahr / PrüfplanKurz:** dauerhaft ausgeblendet über `setshowin*form(false)` (nicht über Field-MERGE `ConditionalShowFormula`); Title nicht pflichtig; Column Formatting zeigt Nachname/Vorname; Wizard-Button **Titel nachziehen** setzt Title + PrüfplanKurz auf bestehende Items.  
4. **Defaults:** Schulform / Terminjahr wo möglich per Field-Patch.

### 5.3 Ansichten & Column Formatting

| Ansicht | Gruppierung / Filter | Zweck |
|---------|----------------------|--------|
| **Nach Klasse** | GroupBy `Klasse`, Sort `Nachname` | Klassenlehrer / Verwaltung |
| **Nach Variante** | GroupBy `Variante`, Sort `Nachname` | Prüfungsorganisation |

Column Formatting (JSON, Setup): farbige Pills für `Variante` (1–3), `Klasse`, `WahlfachMuendlich` (u. a. Seminar/Recht/Geschichte).

---

## 6. Erweiterbarkeit Schulformen

```text
profiles/
  hak.js      ← MVP: Varianten 1–3, Spalten, Defaults Wahlfach
  htl.js      ← später
  …
```

Jedes Profil exportiert:

- `id`, `label`
- `variants[]` (Key, Label, schriftlich[], mündlich[])
- `columns[]` (Graph-Spaltendefinitionen + Form-Regeln)
- `defaultWahlfaecher[]`, `defaultViews[]`
- `buildPruefplanKurz(variante)`

Setup-Wizard wählt Profil → Provisioning bleibt generisch.

---

## 7. Tool-Oberfläche (Setup-Wizard)

Schritte in `sharepoint-liste-srdp.html` (6):

1. **Website & Jahr** – Site-URL, Schulform, Terminjahr  
2. **Klassen** – Abschlussklassen aus Stammdaten (+ optional M365-Gruppen)  
3. **Optionen** – LFS / Wahlfach / Seminar  
4. **Abschlussarbeiten (ABA)** – Liste anlegen, ABA-Excel hochladen  
5. **Rechte** – Vollzugriff / Lesende Rechte / Kandidat:innen (Gruppen + Personen)  
6. **Anlegen** – Anmelde- + Abschlussarbeiten-Liste, Health-Check, Rechte nachziehen  

Zusätzlich: fehlende Spalten ergänzen (idempotent innerhalb desselben Jahres).

---

## 8. Rechte (MVP: beim Setup setzen)

Beim Listen-Anlegen: Vererbung brechen und SharePoint-/Entra-Gruppen zuweisen.

| Stufe (UI) | SPO-Recht | Wer typisch |
|------------|-----------|-------------|
| **Vollzugriff** | Full Control | Direktion, IT, Verwaltung |
| **Lesende Rechte** | Gestaltung (sieht alle Anmeldungen trotz „nur eigene Elemente“) | Lehrkräfte, Sekretariat, Einzelpersonen |
| **Kandidat:innen** | Mitwirken; nur eigene Zeile | Abschlussklassen / Jahrgangsgruppen |

Wizard-Schritt „Rechte“: je Stufe mehrere Entra-/M365-Gruppen und optional Einzelpersonen. Bestehende Config-Keys (`admin*` / `lehrer*`) bleiben intern kompatibel.

---

## 9. Abgrenzung

| Drin (MVP / Phase 4) | Nicht Freitag / später |
|----------------------|------------------------|
| Listen-Provisioning + Wizard | Microsoft Forms |
| **App** Schüler/Lehrer/Admin | Kompensationsprüfungen, Klausurtermine |
| Varianten + Seminar-Bedingung | srdp.at-Anbindung |
| Profile HAK (+ weitere im Setup) | Mehrjährige Sammelliste |
| Jahresweise Liste + Gruppenrechte | |

---

## 10. Phasen

### Phase 0 – Spezifikation (dieses Dokument)

- [x] Formular = SharePoint-Liste  
- [x] Stammdaten + Wizard für Wahlfach/Seminar  
- [x] Jährlich neue Liste  
- [x] HAK-Varianten referenziert  
- [x] Nur Abschlussjahrgänge; Lehrer LFS Pflicht; Seminar Choice+Freitext; Title auto; Rechte setzen  
- [ ] Review Spaltennamen mit Pilot-HAK (optional)  
- [x] Konkrete Gruppennamen / Erkennung Abschlussjahr (`year`-Feld) an Stammdaten-Konvention anbinden  

### Phase 1 – Schema & Provisioning

- [x] `srdp-anmeldung-schema.js` (HAK-Profil + Graph-Spalten)  
- [x] `srdp-anmeldung-logic.js` (Variante → Prüfplan, Seminar-Pflicht, Title-Bau, reine Funktionen + Vitest)  
- [x] `sharepoint-liste-srdp.js` + `tools/sharepoint-liste-srdp.html` Wizard  
- [x] Ansichten anlegen via SPO REST  
- [x] Berechtigungen setzen (Broken Inheritance + Gruppen)  
- [x] Dashboard-Kachel + Hilfe  

### Phase 2 – Formular-Feinschliff

- [x] Conditional Formulas / Pflicht Seminar (Show-Formel)  
- [x] Column Formatting Farben (Variante, Klasse, Wahlfach)  
- [x] Title: Formular ausgeblendet + Anzeige aus Nachname/Vorname + „Titel nachziehen“  
- [x] ClientFormCustomFormatter (Header Prüfplan, Sektionen)  

### Phase 3 – Weitere Schulformen

- [x] Profile HTL, HLW, BAFEB, AHS (`srdp-anmeldung-profiles.js`)  
- [x] Wizard-Schulformwahl + Listenname mit Code  
- [x] Spalten/Formular/Prüfplan pro Profil  
- [x] Tests für alle Profile  

### Phase 4 – Online-App (wie Schulaktivitäten) – Freitag-P0

- [x] `tools/srdp-anmeldung.html` + Shell (`srdp-shell` / linke Nav, Theme wie AKT)  
- [x] Rollen: Schüler · Lehrer · Admin (Demo-Umschalter `?demoRole=1`)  
- [x] Schüler: Formular + Validierung + Zusammenfassung nach Submit  
- [x] Lehrer: Meine Prüfungen + Meine Klasse (KV aus Stammdaten)  
- [x] Admin: Dashboard-KPIs, Liste, Auswertung  
- [x] Graph-CRUD auf bestehender Jahresliste + lokaler Demo-Modus  
- [x] Dashboard-Kachel → App; Setup bleibt unter „Einrichtung“ / `sharepoint-liste-srdp.html`  
- [x] Docs + Hilfe + Workshop-Script Tool C anpassen  
- [x] Bestätigungsmail-Flow (Shared Mailbox) + Setup-ZIP `srdp-mail-setup.html`  

**Feinschliff danach:** Entra-Gruppen strikt, CSV-Export, Landing-Screenshot.

### Phase 5 – Varianten-Wizard + ABA-Abschlussarbeiten (2026-03)

- [x] Profilgesteuerter Wizard (Varianten-Karten → Gebiete → Multi-Lehrer nur mündlich)
- [x] HAK `LehrerAM`; Lehrer-Choice mit Freitext für Mehrfachnamen
- [x] ABA-Portal-Excel-Import (Systemadministration) → Liste `sRDP-Abschlussarbeiten {CODE} {Jahr}`
- [x] Prefill Diplomarbeit / ABA-Titel + Betreuer beim Anmelden
- [x] Tests: Slots, Validierung, ABA-Parse, Wizard-Render

**Power Automate (Bestätigungsmail):** `tools/srdp-mail-setup.html` + Vorlage `assets/power-automate/srdp-anmeldung/` – Trigger neue Anmeldung → Mail aus freigegebenem Postfach mit Zusammenfassung (Variante, Prüfplan, Betreuung, LFS/BKO/Wahlfach). Optional BCC Verwaltung.

**Nicht Freitag:** Klausurplanung, srdp.at-Anbindung, Mehrjahres-Sammelliste.

---

## 11. Festlegungen (ehemals offen)

| Punkt | Entscheidung |
|-------|--------------|
| Klassenfilter | nur **Abschlussjahrgänge** |
| LehrerIn LFS | **Pflicht** |
| Seminar | Choice-Liste **mit Freitexteingabe** (`allowTextEntry`) |
| Item-Titel | **automatisch** `Nachname, Vorname` |
| Rechte | beim Setup **Gruppen setzen** (nicht nur dokumentieren) |

Noch technisch zu klären bei Umsetzung (kein Produkt-Blocker):

- Wie Abschlussjahr in Stammdaten erkannt wird (`classes[].year`, Schulform HAK → 5, …).  
- Ob Title per Column Formatting, berechneter Default oder kurzem Graph-Patch nach Create gesetzt wird (Listenformular hat kein klassisches „berechnet beim Speichern“ wie Access).  
- Welche bestehenden Entra-/SPO-Gruppen die Schule schon hat vs. neu anlegen.

---

## 12. Fertig-Kriterium MVP

Auf der Intranet-Site existiert `sRDP-Anmeldungen {Jahr}` mit allen Spalten; Formular mit Header/Sektionen und Seminar-Show-Formel; Title ausgeblendet mit Nachname/Vorname-Anzeige; farbige Pills; Ansichten „Nach Klasse“ und „Nach Variante“; Berechtigungen gesetzt; Wizard kann erneut laufen und „Titel nachziehen“.
