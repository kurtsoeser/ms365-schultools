# IT-Berechtigungszentrale — Inventar und Plan

**Stand:** 2026-11-20  
**Status:** Welle 1 lesend umgesetzt (`tools/it-berechtigungen.html`). Rückschreiben bleibt offen. Der Kopf jeder Schul-App zeigt `App · Rolle · Quelle` (`src/shared/app-role-chrome.js`).  
**Auftrag:** Ein Ort für die Schul-IT, an dem Berechtigungen aller Schul-Apps sichtbar sind und später in die bestehenden Tool-Einstellungen zurückgeschrieben werden.

Quellen: `src/shared/dashboard-audience-catalog.js`, `src/shared/license-feature-catalog.js`, `src/shared/school-tool-sites.js`, `src/shared/school-audience-groups.js`, `src/shared/school-admin-groups.js`, `src/shared/ablauf-admin-entra.js`, `src/shared/betrieb-admin-entra.js`, die `*-entra-role.js` / `*-permissions.js` / `*-schema.js` der Apps, [`rollen-vertrag.md`](./rollen-vertrag.md), [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md).

---

## 1. Kurzbefund

Im Werkzeugkatalog stehen **100** Einträge. Davon sind **27** im Portal für Lehrkräfte, Schülerinnen und Schüler oder alle sichtbar. Die übrigen **73** sind Schul-IT: Mitgliedschaften, Hygiene, Schuljahresstart, Playbooks, Power-Automate-Rezepte. Die haben keine eigene „wer genehmigt / wer reicht ein“-Matrix. Sie bleiben im Katalog.

Die Zerfahrenheit sitzt bei den **Alltags-Apps**. Jede hat über die Jahre eigene Gruppen, Einzelpersonen, Postfächer und SharePoint-Listen bekommen. Ein Teil liest schon gemeinsame Sammelgruppen (Lehrer, Schüler, Schulleitung, Verwaltung). Ein anderer Teil hält eine eigene Kopie in `localStorage`. Die SharePoint-Listenrechte sind ein dritter, davon getrennter Schreibweg.

Eine Zentrale lohnt sich. Sie darf kein Formular mit 100 Werkzeugen werden. Sie setzt **wenige Schul-Prinzipale** und mappt sie auf **Slots** je App. Domäne (Schuljahr, Varianten, Buchungen, Wahlphasen) bleibt im Tool.

---

## 2. Drei Schichten bleiben getrennt

| Schicht | Frage | Heute | Zentrale |
|--------|--------|--------|----------|
| Lizenz | Hat die Schule das Modul? | `license-feature-catalog` | unverändert |
| Portal-Persona | Wer bin ich im Portal? | `it` · `lehrer` · `schueler` | unverändert |
| App-Rolle | Was darf ich in diesem Tool? | je App anders | das ist der Gegenstand |

UI-Rollen ersetzen keine SharePoint-Rechte. Eine Zentrale muss beides kennen: **wer die App-Rolle bekommt** und **welches Listenrecht** die Gruppe auf welcher Liste hat.

---

## 3. Was schon zentral ist

| Baustein | Datei | Inhalt |
|----------|--------|--------|
| Lehrer- und Schüler-Sammelgruppe | `school-audience-groups.js` | aus Stammdaten-Setup (`matched.lehrerGroupId` / `schuelerGroupId`), Fallback `ms365-dashboard-audience-groups-v1` |
| Schulleitung und Verwaltung | `school-admin-groups.js` | `schulleitungGroupId`, `verwaltungGroupId`; füllt leere Admin-/Direktions-Slots |
| Ablauf-Admin | `ablauf-admin-entra.js` | Schulleitung, Verwaltung, Schularbeiten-Admin, Freistellungs-Direktion, Schul-IT, plus Legacy-Mail-Listen |
| Betrieb-Admin | `betrieb-admin-entra.js` | derselbe Helfer für Aufgaben, Technik, Raumplan, Ausgaben |
| Site-URL | `school-tool-sites.js` | 16 Tool-Sites, Auflösung Override → `planerSiteUrls` → Default → Intranet |
| KV | Stammdaten Klassen | `headEmail` / `KVEmail`, keine eigene Portal-Persona |
| Schul-IT | `designated-school-it.js` | Global Admin, designierter IT-Kontakt, Betreiber |

Lücke: Die Sammelgruppen sind **Fallback, wenn das Tool-Feld leer ist**. Viele Tools speichern danach eine eigene Kopie. Ablauf und Betrieb haben **keine eigenen Gruppen-Picker**; Genehmigungs-Planer schon.

---

## 4. Inventar der Schul-Apps mit Rechte-Slots

Speicherorte sind Browser-`localStorage` auf dem Rechner, der das Setup ausgeführt hat. Sie sind kein schulweites System of Record.

### 4.1 Genehmigung (`approval-planer`)

| App | App-Rollen | Was die IT setzt | Listen | Speicher |
|-----|------------|------------------|--------|----------|
| Schulaktivitäten | `admin`, `kv`, `lehrer` | Gruppe + Personen: Direktion/Admin, Lehrkräfte. KV kommt aus Klassen-Stammdaten (`headEmail`). Flow: Genehmiger-Mail, Kalenderkonten | `Schulaktivitaeten` (Lehrer: Mitwirken, Admin: Vollzugriff), `Aktivitaet-Regelwerk` (Lehrer: Lesen), Stammdaten Lehrer/Klassen lesen | `ms365-schulaktivitaeten-perms-v1` |
| Freistellung Schüler | `admin`, `direktion`, `kv`, `schueler` | Grant-Matrix: Gruppe oder Person × Rolle. Jahrgang, Klassenkatalog. Genehmiger-Mail | `Freistellungen` | `ms365-freistellung-perms-v1`, `ms365-freistellung-setup-v1` |
| Freistellung Lehrkräfte | `direktion`, `lehrer` | Gruppe + Personen: Lehrkräfte, Direktion. Genehmiger-Mail, Fallback auf die Schüler-Freistellung | `Lehrer-Freistellungen` | `ms365-lfr-perms-v1`, `ms365-lfr-setup-v1` |

Setup-Geschwister (nur IT): `schulaktivitaeten-setup`, `freistellung-setup`, `lehrer-freistellung-setup`, dazu die Listen-Assistenten. Dieselbe Schale (`it-setup-shell`), andere Domänen-Schritte. `lehrer-freistellung-setup` ist im Katalog noch nicht als eigene Kachel geführt.

### 4.2 Register (`register-planer`)

| App | App-Rollen | Was die IT setzt | Listen | Speicher |
|-----|------------|------------------|--------|----------|
| Schularbeiten | `admin`, `lehrer` (schreibt), `schueler` (liest) | Drei Gruppen + Einzelpersonen. Schüler auf der Terminliste nur Lesen | `SAP-Schularbeiten`, `SAP-Regelwerk`, `SAP-Terminfenster`, `SAP-FachMeta` | `ms365-schularbeiten-perms-v1` |

Diese Konfiguration ist heute die **Referenz**, die andere Apps mitlesen: Projektwochen, Wahlfach-Admin, Ablauf- und Betrieb-Admin.

### 4.3 Buchung (`projekt-buchung`)

| App | App-Rollen | Was die IT setzt | Listen | Speicher |
|-----|------------|------------------|--------|----------|
| Projektwochen | `admin`, `lehrer`, `schueler` | Keine eigene Matrix. Liest Schularbeiten-Gruppen und Stammdaten-Sammelgruppen. Zusätzlich Bookings-Export | `PW-Aktionen`, `PW-Angebote` | Site `ms365-pw-site-url`; Rollen aus SA |
| Elternsprechtag | Lehrkraft | Buchungskonto / Bookings, kein Genehmiger-Modell | — | Playbook + Tool, nicht die Planer-Matrix |

### 4.4 Wahl und Anmeldung (`wahl-portal`)

| App | App-Rollen | Was die IT setzt | Listen | Speicher |
|-----|------------|------------------|--------|----------|
| sRDP-Anmeldung | `schueler`, `lehrer`, `admin` | Eigene Listenrechte, **mehrere** Gruppen je Slot: Admin/Direktion (Vollzugriff), Lehrkräfte (Lesen), Kandidat:innen (Mitwirken) plus Einzelpersonen. App-Rolle selbst ist noch Heuristik aus Stammdaten, nicht dieselbe Matrix | `sRDP-Anmeldungen {Jahr}` bzw. `sRP-Anmeldungen {Jahr}` | `ms365-srdp-list-perms-v1`, App-State `ms365-srdp-anmeldung-v1` |
| Wahlfach-Zuteilung | Admin (IT/SLG), Schüler auf eigener Anmeldeseite | Zugang über Schularbeiten-Admin-Gruppe. Kein eigener Picker | `Wahlfach-Runden`, `Wahlfach-Angebote`, `Wahlfach-Anmeldungen`, `Wahlfach-Zuteilungen` | `ms365-wahlfach-zuteilung-settings-v1` (darin Site-URL) |
| Schulsprecherwahl | Wahlleitung vs. Wählende | `isAdmin` kommt vom Wahl-Dienst, nicht aus dem Planer-Gruppenpicker. Zwei Arten: Schulsprecher, SGA-Lehrer | eigener Dienst, keine SPO-Planer-Liste | späterer Slot „Wahlleitung“, nicht in Welle 1 |

### 4.5 Ablauf (`ablauf-kommunikation`)

Gemeinsam: Absender ist Lehrkraft oder Leitung. Empfänger antworten über Link oder Formular. Admin = Schulleitung / Verwaltung / SA-Admin / Schul-IT (`ablauf-admin-entra.js`). KV, wo nötig, aus Klassen-`headEmail`.

| App | Rollen in der App | Zusätzlich im Tool | Listen |
|-----|-------------------|--------------------|--------|
| Einsammler | Admin vs. Lehrkraft | Legacy `adminEmails` | `Einsammler`, `EinsammlerAntworten` |
| Elternbrief | Admin vs. Lehrkraft | Absender: eigenes Konto oder zentrales Postfach; Legacy `adminEmails` | `Elternbriefe`, `ElternbriefLesebestaetigungen` |
| Krankmeldung | Sekretariat/Admin vs. KV; öffentliche Meldung | kein eigener Gruppenpicker | `Krankmeldungen` |
| Kollegiumsumfragen | Admin, Lehrkraft, KV | Zielgruppe je Umfrage: alle Lehrkräfte, KV, Abteilung, manuell | `KollegiumUmfragen`, `KollegiumUmfrageAntworten` |

Site-URL liegt jeweils in `ms365-<tool>-sp-settings-v1` bzw. bei Elternbrief zusätzlich `ms365-elternbrief-settings-v1`.

### 4.6 Betrieb

Gemeinsam: Melder vs. Admin über `betrieb-admin-entra.js` (dieselben Gruppen wie Ablauf).

| App | Rollen | Zusätzlich | Listen |
|-----|--------|------------|--------|
| Aufgaben | `melder`, `admin` | Zuweisung am Datensatz: Direktion, IT, KV, Lehrer, Verwaltung | `Aufgaben` |
| Technik-Meldungen | `melder`, `admin` | Kategorien | `TechnikMeldungen` |
| Raumplan | Schreiben nur Admin, sonst lesen | — | `Raeume` |
| Ausgaben | Admin vs. „meine Vorgänge“ | Legacy `adminEmails` | `AusgabenGegenstaende`, `AusgabenVorgaenge` |

### 4.7 Kalender und Team-Anträge

| App | Was die IT setzt | Listen / Konten |
|-----|------------------|-----------------|
| Kalender und Termine | Postfächer Schul / Lehrer / Schüler. Lesen und Schreiben je Persona | Liste `SchulTermineHub` (Alt: `Schultermine`). Schlüssel `ms365-termine-hub-settings-v1`, Konten aus Kalender-Postfächer |
| Team-Antrag | Freigeber-Gruppe und Freigeber-Mails | `TeamAntraege` |

### 4.8 Gemeinsame Stammdaten-Listen

Nicht eine App, aber die Unterlage für KV und Lehrkräfte:

| Listen | Rechte-Profil | Speicher |
|--------|----------------|----------|
| `Lehrerinnen`, `Klassen` (Titel aus Intranet-Setup) | Admin Vollzugriff, Lehrer lesen, Schüler je Profil | `ms365-stammdaten-listen-perms-v1` |

Schüler-Stammdaten-Portal ist ein Sonderfall: Schul-IT und Sekretariat, Eltern mit PIN, Schreiben über Gateway. Kein Genehmigungs-Planer.

Diplomarbeiten hat nur Betriebsmodus (Kanäle vs. Legacy) und Typ (Diplom/VWA). Keine Rollenmatrix.

---

## 5. Slot-Typen, die sich wiederholen

Dieselben Bausteine, verschiedene Namen:

| Slot | Bedeutung | Kommt vor als |
|------|-----------|----------------|
| Leitung | darf genehmigen, Regeln, Vollzugriff | `admin`, `direktion`, Schulleitung, Freigeber |
| Lehrkräfte | reichen ein oder tragen ein, Mitwirken oder Lesen je Liste | `lehrer`, `groupLehrer` |
| Schüler / Kandidat:innen | eigene Anträge oder nur lesen | `schueler`, `kandidaten` |
| Klassenvorstand | eigene Klasse, Stellungnahme, Kenntnis | `kv`, abgeleitet aus `headEmail` |
| Verwaltung / Sekretariat | Inbox, Ausgaben, Krankmeldung | Verwaltung-Sammelgruppe, teils Mail-Liste |
| Technik | Tickets bearbeiten | Betrieb-Admin, oft dieselbe IT-/Leitungsmenge |
| Konto | Flow-Genehmiger, Absender-Postfach, Kalender, Bookings | E-Mail, kein Gruppenrecht |
| Einzelperson | zusätzlich zur Gruppe | `*Users` in den Planer-Matrizen |

Listenrecht ist **nicht** dasselbe wie App-Rolle. Typisch: Leitung `fullControl`, Lehrkraft auf der Arbeitsliste `contribute` und auf Regelwerk `read`, Schüler `read` oder `null`.

---

## 6. Was in der Zentrale nichts zu suchen hat

- Die 73 IT-Werkzeuge (Gruppen anlegen, Hygiene, Import, Playbooks). Sichtbarkeit bleibt „nur Schul-IT“.
- Lizenzpakete.
- Portal-Persona.
- Fachdaten: sRDP-Jahr und Varianten, Schularbeiten-Regelwerk-Inhalt, Wahlphasen, Umfrage-Zielgruppe je Vorgang, Krankmeldungs-Kategorien, Bookings-Inhalte.
- Demo-Rollen (`?demoRole=1`).

---

## 7. Zielbild

Eine IT-Seite, nur Persona Schul-IT.

**Oben: Schul-Prinzipale** (einmal setzen, aus Stammdaten vorbelegen):

1. Lehrkräfte-Gruppe  
2. Schüler-Gruppe  
3. Schulleitung  
4. Verwaltung / Sekretariat  
5. Schul-IT  
6. optionale Einzelpersonen je Principal  
7. Konten: Direktion-Genehmiger, zentrales Eltern-Postfach, Schulkalender, Lehrer- und Schülerkalender  

KV bleibt abgeleitet. Keine siebte Sammelgruppe erfinden.

**Darunter: eine Zeile je Schul-App**, gefiltert nach Familie.

Pro Zeile nur die Slots, die die App wirklich hat. Standard ist „übernehmen“. Abweichung ist möglich (sRDP-Kandidat:innen sind nicht die ganze Schüler-Gruppe; Lehrer-Freistellung hat keine Schüler-Rolle).

**Rückschreiben** geht an die bestehenden Schlüssel und, wo die App das schon kann, an die SharePoint-Listenrechte. Die Apps lesen weiter ihre eigenen Loader (`loadAktPermissionsConfig`, `loadLfrPermissions`, …). Die Zentrale ist ein Schreiber davor, kein zweites Rechtesystem.

**Sicherung (umgesetzt, 2026-11-19):** Beim Sichern in die IT-Sicherungsbibliothek schreibt das Konfigurationspaket die Rechte- und Setup-Schlüssel mit. Ein späteres Laden auf einem anderen Gerät übernimmt sie. Das ist kein Live-Abgleich. Enthalten sind Schulaktivitäten (Gruppen, Setup, Site, Flow-Konten), sRDP (Listenrechte, Einrichtung, Mail-Setup, Site) sowie die schon zuvor gesicherten Planer (Schularbeiten, beide Freistellungen) und die Ablauf-/Betriebs-Einstellungen. Die eigentlichen Listeneinträge bleiben in SharePoint, nicht in dieser Sicherung.

---

## 8. Wellen (wenn wir bauen)

| Welle | Inhalt | Ergebnis |
|-------|--------|----------|
| 0 | Dieses Inventar. Kein Code | Plan steht |
| 1 | Nur lesen: eine Seite zeigt je App die heutigen Gruppen, Mails, Listen, Speicherort. Nichts schreiben | IT sieht die Zerfahrenheit an einem Ort |
| 2 | Prinzipale als Quelle. Leere Tool-Slots füllen sich daraus. Bestehende Overrides bleiben | Lehrer/Schulleitung nicht mehr fünfmal neu wählen |
| 3 | Rückschreiben Genehmigung + Schularbeiten: Schulaktivitäten, beide Freistellungen, Schularbeiten. Inkl. Listenrechte, die diese Apps schon setzen | Die lautesten Doppel-Picker fallen weg |
| 4 | Ablauf und Betrieb: Legacy-`adminEmails` nur noch Ausnahme. Elternbrief-Postfach und Termine-Postfächer als Konten-Slots | Eine Leitung, eine Verwaltung |
| 5 | Sonderfälle: sRDP mehrere Gruppen je Slot, Team-Antrag-Freigeber, Wahlfach bleibt SA-Admin, Schulsprecherwahl nur wenn der Dienst eine Gruppe kennt | Kein erzwungenes Einheitsformular |

Jede Welle: die App-Seite zeigt denselben Stand wie die Zentrale. Ein gesetzter Wert darf nicht nur in der Zentrale sichtbar sein.

---

## 9. Nicht tun

- 100 Kacheln in ein Berechtigungsformular kopieren.
- App-Rollen auf die drei Portal-Personas zusammenziehen.
- KV als Entra-Gruppe erzwingen.
- Bookings, Wahlphasen oder sRDP-Varianten in die Zentrale ziehen.
- SharePoint-Rechte nur in der UI ändern und den Graph-Schreibweg weglassen.
- Eigene zweite Speicherung, die die Tools nicht lesen.

---

## 10. Offene Punkte vor Welle 1

1. Ablage: Tenant-Einstellungen der Schule, oder eine SharePoint-Liste auf der Intranet-Site?
2. Wer darf die Zentrale öffnen: nur designierte Schul-IT, oder auch Schulleitung nur lesen?
3. Schüler-Freistellung-Grant-Matrix ist feiner als „eine Gruppe = eine Rolle“. Welle 3 muss mehrere Rollen pro Gruppe behalten.
4. sRDP-App-Rolle und sRDP-Listenrecht sind heute zwei Mechanismen. Die Zentrale zeigt beide, führt sie nicht still zusammen.
5. Katalog-Lücke: `lehrer-freistellung-setup` als IT-Kachel, sobald die Seite produktiv ist.
