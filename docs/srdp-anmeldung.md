# sRDP-Anmeldung – Umsetzungsplan

**Stand:** 2026-09-23  
**Status:** Planung (noch kein Code)  
**Zielgruppe MVP:** HAK (Tagesform / AUL analog, soweit gleiche Varianten)  
**UI:** Vanilla JS/ESM, Look an `app.css` – Setup-Wizard + SharePoint-Listenformular  

Verwandt: Schularbeiten-Listen-Setup, Intranet-Hub, Stammdaten (`tenant-settings` / `app-data-v2`), PA-Rezept-Muster nur als Referenz (hier **kein** Microsoft Forms).

Quellen Varianten-Logik: [hak.cc – RDP](https://www.hak.cc/pruefungen-abschluss/rdp), BMB sRDP BHS, Schüler:innen-Infos HAK; srdp.at deckt vor allem die **zentralen** Klausurfächer (D, LFS, AM) ab, nicht die HAK-Variantenmatrix.

---

## 1. Zielbild

Admin richtet **pro Haupttermin / Jahr** eine SharePoint-Liste ein. Schüler:innen melden sich über das **native SharePoint-Listenformular** an. Direktion / Verwaltung arbeitet mit vorbereiteten **Ansichten** (nach Klasse, nach Variante, …).

```text
Admin ──MSAL──► sRDP-Anmeldung Setup (ms365-schultools)
                      │ Graph Sites.ReadWrite.All
                      ▼
             SharePoint Intranet-Site
             └─ sRDP-Anmeldungen {Jahr}
                      │
Schüler:innen ───────► Listenformular (Neues Element)
Verwaltung ───────────► Ansichten (Klasse / Variante / …)
```

**System of Record:** eine neue Liste pro Terminjahr (archivierbar, kein Vermischen der Jahrgänge).

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

Implikation für die Anmeldung:

- **LFS** kommt in **allen** Varianten vor (schriftlich oder mündlich) → Feld `LFS` (+ optional `LehrerIn LFS`) bleibt grundsätzlich sichtbar.
- **Wahlfach** + `LehrerIn Wahlfach` immer.
- **BKO** + `LehrerIn BKO` immer.
- **AM** braucht kein eigenes Anmelde-Dropdown (kein Sprachcode); die Variante legt fest, ob AM schriftlich oder mündlich ist → als **abgeleitete/info**-Spalten oder nur in der Hilfe/Ansicht.
- Wenn Wahlfach = „Seminar …“ → Feld `Seminar` Pflicht (Bezeichnung des Seminars).

Offizielle Wahlfach-Bezeichnungen (Wizard-Defaults, Schule kann kürzen/erweitern): Religion/Ethik, Kultur, Geschichte und …, Geografie und …, Naturwissenschaften, Recht, Volkswirtschaft, Berufsbezogene Kommunikation in der LFS, Mehrsprachigkeit, Wirtschaftsinformatik, Seminar …, Freigegenstand …, (Sonderfälle Slowenisch/Deutsch nur zweisprachige HAK).

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
| `LehrerLFS` | LehrerIn LFS | Choice | **ja** | Pflicht |
| `LehrerBKO` | LehrerIn BKO mündlich | Choice | ja | |
| `WahlfachMuendlich` | Wahlfach mündlich | Choice | ja | |
| `Seminar` | Seminar | Choice **mit Freitext** (`allowTextEntry: true`) | bedingt | Pflicht wenn Wahlfach Seminar; Optionen + eigene Eingabe |
| `LehrerWahlfach` | LehrerIn Wahlfach | Choice | ja | |
| `Bestaetigung` | Bestätigung Anmeldung | Boolean / Choice Ja | ja | |
| `Schulform` | Schulform | Choice (hidden) | – | Default `HAK` für spätere Profile |
| `TerminJahr` | Terminjahr | Zahl/Text | – | z. B. 2026 |
| optional `PruefplanKurz` | Prüfplan (Info) | Text berechnet/Beschreibung | – | z. B. „schriftlich: D, BFK, LFS · mündlich: BKO, AM, Wahlfach“ |

Referenzen als **Choice-Text** (Namen/Codes), keine Person-Felder im MVP – entspricht den Screenshots mit Pills.

### 5.2 Formular-Logik (Listenformular)

SharePoint Modern Form – per Setup gesetzt:

1. **ClientFormCustomFormatter:** Header mit Live-Name (`Nachname, Vorname`) und Prüfplan zur Variante; Body in Sektionen; Footer-Hinweis.  
2. **Seminar:** `ConditionalShowFormula` – sichtbar wenn `WahlfachMuendlich` mit „Seminar“ beginnt.  
3. **Title:** im Formular ausgeblendet (`=false`), nicht pflichtig; Column Formatting zeigt Nachname/Vorname; Wizard-Button **Titel nachziehen** setzt Title + PrüfplanKurz auf bestehende Items.  
4. **Schulform / Terminjahr / PrüfplanKurz:** im Formular ausgeblendet; Defaults wo möglich.

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

Schritte (Vorschlag):

1. **Site & Jahr** – Intranet-URL, Terminjahr, Listenname-Vorschau  
2. **Schulform** – HAK (weitere disabled / „demnächst“)  
3. **Klassen** – Multi-Select aus Stammdaten (vorfilter Abschlussjahr)  
4. **Lehrer-Pools** – alle Lehrer oder gefilterte Listen für die vier Choice-Felder  
5. **LFS / Wahlfach / Seminar** – Optionen bearbeiten  
6. **Gruppen & Rechte** – Admin-, Lehrer-, Kandidaten-Gruppen wählen/anlegen; Vererbung brechen  
7. **Anlegen** – Liste + Spalten + Ansichten + Formatting + Rechte; Protokoll  
8. **Fertig** – Link zur Liste, Link „Neues Element“

Zusätzlich: **Health-Check** (Liste existiert, Spalten ok, Ansichten ok) und „fehlende Spalten ergänzen“ (idempotent innerhalb desselben Jahres).

Integration: eigene Tool-Seite + optional Checkbox im Intranet-Hub-Startpaket (wie Schularbeiten).

---

## 8. Rechte (MVP: beim Setup setzen)

Beim Listen-Anlegen: Vererbung brechen und SharePoint-/Entra-Gruppen zuweisen.

| Rolle | Recht | Gruppe (Vorschlag, im Wizard konfigurierbar) |
|-------|--------|-----------------------------------------------|
| Schüler:innen Abschlussjahrgänge | Beiträge hinzufügen; eigene Items lesen/bearbeiten | z. B. Jahrgangs-/Klassengruppen oder `sRDP-Kandidaten-{Jahr}` |
| Verwaltung / Direktion | Vollzugriff | z. B. `sRDP-Admin` / bestehende Admin-Gruppe |
| Klassenlehrer / Lehrkräfte | Lesen | z. B. Lehrer-Gesamtgruppe oder `sRDP-Lehrer` |

Wizard-Schritt „Gruppen“: vorhandene Gruppen wählen oder Namen vorgeben; Setup setzt Berechtigungen und protokolliert das Ergebnis. Fallback-Hinweis, falls Graph die Rechte nicht setzen kann.

---

## 9. Abgrenzung

| Drin (MVP) | Nicht im MVP |
|------------|--------------|
| Listen-Provisioning HAK | Microsoft Forms / Power Automate |
| Wizard Stammdaten + manuelle Choices | Eigene Anmelde-UI in den Schultools |
| Varianten-Infos + Seminar-Bedingung | Kompensationsprüfungen, Terminplanung Klausuren |
| Ansichten Klasse + Variante | HTL/andere Profile (nur Hook) |
| Jahresweise neue Liste + Gruppenrechte | Mehrjährige Sammelliste |

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
