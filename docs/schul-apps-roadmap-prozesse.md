# Schul-Apps Roadmap: Prozesse, Einsammler & Ausgaben

**Stand:** 2026-11-06  
**Status:** P0–P5 + Elternbrief; neu: Teams-Antrag (`docs/teams-antrag.md`), Diplom/VWA Umbau C (`docs/diplomarbeiten.md`)  
**Kontext:** Ergänzung zu Lizenzpaketen (`license-modules.md`) und Marktanalyse (Untis/WebUntis, edu.FLOW, Klasso, Schulmanager).  
**Weiterführend:** Lehrer-End-User-Portal, Kalender und Termine, Aufgaben, Technik, Raumplan → [`lehrer-portal-roadmap.md`](./lehrer-portal-roadmap.md)  
**Nicht im Scope:** Zahlungen / Payment-Gateways (bewusst ausgeschlossen).  
**Tool-Doku:** [`einsammler.md`](./einsammler.md), [`elternbrief.md`](./elternbrief.md)

---

## 0. Zielbild in einem Satz

Neue Schul-Apps sollen **wiederkehrende Papier-/Excel-Workflows** im Schulalltag ersetzen – mit dem gleichen Bauplan wie Schularbeiten-/Schulaktivitäten-Planer: **SharePoint als System of Record**, schlanke UI im Tenant, optional öffentliches Formular für Eltern, Lizenz über Pakete.

---

## 1. Plattform-Muster (damit später viele Apps dazukommen)

Alles Neue folgt demselben Gerüst. Neue Tools = neue **Schema-Datei + Setup + UI**, nicht jedes Mal neue Architektur.

### 1.1 Schichten

```text
┌─────────────────────────────────────────────────────────────┐
│  Lizenz (feature-catalog)  →  Paket schulapps / neu „ablauf“ │
├─────────────────────────────────────────────────────────────┤
│  Dashboard-Kachel (data-tool-id) + optional Playbook/Setup  │
├─────────────────────────────────────────────────────────────┤
│  Tool-UI (Vanilla/ESM, app.css)  │  optional Public Form     │
│  MSAL Lehrkraft/Admin            │  Token/Selfhost wie       │
│                                  │  Stammdaten-Portal        │
├─────────────────────────────────────────────────────────────┤
│  Schema (*-schema.js): Listen, Spalten, Choice-Sets, Status │
├─────────────────────────────────────────────────────────────┤
│  SharePoint Intranet-Site (Graph Sites.ReadWrite.All)       │
│  Listen + Views + optional Dokumentbibliothek (PDFs/ICS)    │
├─────────────────────────────────────────────────────────────┤
│  optional Power Automate (Mail/Reminder) – Setup-ZIP        │
└─────────────────────────────────────────────────────────────┘
```

### 1.2 Zwei Archetypen (alles Weitere ist Variation)

| Archetyp | Kernfrage | Beispiele |
|----------|-----------|-----------|
| **A – Einsammeln** | „Wer hat geantwortet / zugestimmt / gelesen?“ | Einwilligung, Elternbrief, Krankmeldung, Kollegiumsumfrage |
| **B – Ausgabe** | „Was wurde an wen ausgegeben – und ist es zurück?“ | Schlüssel, Raumkarte, Parkplatz, Schrankenkarte, Laptop |

Beide speichern in **SharePoint-Listen**, nutzen **Klassen/Lehrer aus Stammdaten**, haben **Statusfluss**, **Filter „meine / Klasse / offen“**, Export (CSV/Excel), optional Benachrichtigung.

### 1.3 Pflicht-Bausteine je neuem Tool (Checkliste)

1. **`toolId`** im `feature-catalog` (+ Dashboard-Kachel + Audience)
2. **`*-schema.js`** – Listen-Titel, Spalten, Choices, Defaults
3. **`sharepoint-liste-*.js` + HTML** – Listen anlegen / Spalten nachziehen
4. **Tool-HTML + UI-Modul** – Liste, Formular, Filter, Rollen (Demo → Entra)
5. **Kurz-Doku** unter `docs/` (wie Schulaktivitäten)
6. **Tests** – Schema/Logic mind. Unit; optional UI-Smoke
7. **Lizenzpreis** – Paket zuordnen (siehe §7)

### 1.4 Gemeinsame Felder (Konvention)

Wiederkehrende Spaltennamen, damit Reports und Flows wiederverwendbar bleiben:

| Intern | Typ | Bedeutung |
|--------|-----|-----------|
| `Title` | text | Anzeige-Titel |
| `VorgangId` | text | stabile ID (`ein-…`, `brk-…`, `aus-…`) |
| `Status` | choice | pro Tool definiert, immer lowercase |
| `KlasseCode` / `KlasseCodes` | text | Stammdaten-Referenz |
| `SchuelerName` / `SchuelerEmail` | text | wo nötig |
| `LehrerCode` / `LehrerEmail` | text | Antragsteller / verantwortlich |
| `FaelligAm` | dateOnly | Frist |
| `ErledigtAm` | dateTime | Abschluss |
| `Notiz` | multiline | Freitext |
| `Quelle` | choice | `ui` \| `public` \| `import` |

### 1.5 Öffentliche Endpunkte (Eltern)

Nur wo nötig (Einwilligung, Elternbrief-Leselink, Krankmeldung):

- Pattern wie **Stammdaten-Portal**: schlankes Public-HTML, Tenant-Config, Speichern in SharePoint über Backend **oder** Graph-App-Only später.
- MVP-Variante: **Magic-Link** (Item-ID + HMAC/Token in Query), kein Eltern-Login.
- Keine Zahlungsfelder, keine IBAN, keine PSP-Anbindung.

### 1.6 Was bewusst gleich bleibt

- Kein zweites „Schul-SIS“ – Untis/Sokrates bleiben Quelle für Plan/Stammdaten.
- Kein Zahlungsmodul.
- Kein volles digitales Klassenbuch / Notenbuch.
- M365-Tenant der Schule = Datenhoheit.

---

## 2. Produktsteckbrief: Digitale Einwilligungen (Einsammler)

### Kurz

Lehrkraft/Admin startet einen **Einsammler** (z. B. Exkursion, Foto/Jahrbuch, Datenschutz). Eltern antworten per Link inkl. Zustimmung. Schule sieht **offen / zugestimmt / abgelehnt / Frist überschritten**.

### Nutzen

Ersetzt Zettelmappe + „Wer fehlt noch?“-Excel. Direkter edu.FLOW-/Klasso-Schmerzpunkt, ohne separates Eltern-Ökosystem.

### Personas

| Rolle | Kann |
|-------|------|
| Lehrkraft / KV | Einsammler für Klasse(n) anlegen, Reminder, Export |
| Admin / Direktion | schulweite Vorlagen, Übersicht alle Klassen |
| Eltern | Link öffnen, lesen, zustimmen/ablehnen, optional Kommentar |

### Funktionsumfang MVP

- Vorlagen: Exkursion, Foto, Datenschutz, Wandertag, Sonstiges (frei)
- Zielgruppe: Klasse(n) oder manuelle Schülerliste
- Felder: Text, Ja/Nein, optionale Unterschrift (Canvas → Bild in Anlage oder Status „unterschrieben“)
- Status je Schüler: `offen` → `zugestimmt` \| `abgelehnt` \| `abgelaufen`
- Dashboard: Fortschrittsbalken, Filter „nur offene“, CSV-Export
- Reminder-Mail (optional PA) an offene Empfänger

### Nicht im Scope (MVP)

- Mehrstufige Genehmigungsketten, PDF-Rechtsgültigkeit qualifizierter Signatur, Zahlungen

### Datenmodell (Kern)

**Liste `Einsammler`** (Kopf) + **Liste `EinsammlerAntworten`** (eine Zeile je Empfänger)

Kopf: Titel, Typ, Text/HTML, Frist, KlasseCodes, Ersteller, Status (`entwurf`/`aktiv`/`geschlossen`)  
Antworten: EinsammlerId, SchuelerName, ElternEmail, Status, AntwortAm, Kommentar

### Lizenz / Preis (Orientierung)

Paket **Schul-Apps** oder neues Add-on **„Eltern & Ablauf“** · **150–190 €**/Schuljahr

### Aufbau-Reihenfolge

1. Schema + Listen-Setup  
2. Interne UI (ohne Public) – manuelles Abhaken  
3. Public Magic-Link Form  
4. Reminder-Flow  

---

## 3. Produktsteckbrief: Elternbrief mit Lesebestätigung

**Umgesetzt:** eigenes Tool `tools/elternbrief.html` (+ Leseseite, SharePoint-Setup). Details: [`elternbrief.md`](./elternbrief.md).

### Kurz

Schule veröffentlicht einen **Elternbrief** (Vorlage, Text, optional Anhang). Pro Familie/Schüler: persönlicher Link + **Lesebestätigung**.

### Nutzen

„Wer hat das gelesen?“ ohne WhatsApp-Klassenchat. Vorlagen und Absender-Zentrale für den Schulalltag.

### Personas

Lehrkraft/KV/Admin verfasst; Eltern lesen und bestätigen; Admin sieht alle Briefe.

### Funktionsumfang MVP

- Vorlagen + Editor, optionaler Anhang
- Ziel: Klasse(n) aus Stammdaten (Eltern-Mails)
- Zustellung: persönlicher Link; mailto als Versandhilfe (kein Auto-Serienmail)
- Tracking: `offen` → `gelesen` (Button „Zur Kenntnis genommen“)
- Übersicht + Fortschritt; Absender eigenes Konto oder zentrales Postfach

### Architektur-Hinweis

**Entscheidung:** eigenes Tool (nicht nur Einsammler-Modus) – Vorlagen, Anhang, Absender-Settings und Brief-Übersicht brauchen eigenen Scope. Einsammler bleibt für Einwilligung/Abgabe.

### Lizenz

Im Paket **`ablauf`** (Eltern & Ablauf).

---

## 4. Produktsteckbrief: Einfache Krankmeldung

**Umgesetzt (P3):** Admin `tools/krankmeldung.html`, Public `tools/krankmeldung-formular.html`, SharePoint-Setup, optional Remote-API. Details: [`krankmeldung.md`](./krankmeldung.md).

### Kurz

Eltern melden Kind **krank** (Datum von–bis, optional Grund-Kategorie). Sekretariat/KV sieht Tagesliste; optional Mail an Klassenvorstand.

### Nutzen

Entlastet Telefon am Morgen. Bewusst **einfach** – kein volles Absenzen-/Klassenbuch-System (das bleibt Untis/WebUntis).

### Personas

Eltern (Public Form); Sekretariat (Inbox); KV (Klassenfilter); optional Lehrkraft „meine Klasse“.

### Funktionsumfang MVP

- Public Form: Schülername, Klasse, von/bis, Kategorie (`krankheit`/`arzt`/`sonstiges`), Kontakt
- Liste `Krankmeldungen`: Status `neu` → `zur_kenntnis` → `archiviert`
- Tages-/Wochenansicht, Filter Klasse
- Button „KV benachrichtigen“ (Graph Mail / PA)
- Duplikat-Hinweis (gleiche Klasse+Name+Tag)

### Nicht im Scope

Unterrichtsstunden-genaue Fehlzeiten, Entschuldigungs-PDF-Workflow, Bindung an WebUntis-Absenzen (optional später Export).

### Datenmodell

Eine Liste reicht. Felder: SchuelerName, KlasseCode, Von, Bis, Kategorie, GemeldetAm, GemeldetVonEmail, Status, KenntnisLehrerEmail, Notiz.

### Lizenz / Preis

Teil von **„Eltern & Ablauf“** oder einzeln **80–110 €**/Schuljahr (Schulmanager Beurlaubung-Niveau).

---

## 5. Produktsteckbrief: Kollegiumsumfragen

**Umgesetzt (P4):** Admin `tools/kollegium-umfragen.html`, MSAL-Antwort `kollegium-umfragen-antwort.html`, SharePoint-Paket, Vorlagen. **SoR = SharePoint** (nicht Forms). Details: [`kollegium-umfragen.md`](./kollegium-umfragen.md).

### Kurz

Direktion/Admin startet Umfragen **nur für Lehrkräfte**: Klassenwünsche, Stundenplanwünsche, SCHILF-Themen, Verfügbarkeit, Ja/Nein-Beschlüsse.

### Nutzen

edu.FLOW-intern ohne Elternkanal. Schnelle Entscheidungen, auswertbar.

### Personas

Admin/Direktion erstellt; Lehrkräfte antworten (MSAL – **kein** Public-Link nötig); optional anonymisierte Auswertung.

### Funktionsumfang MVP

- Fragetypen: Single Choice, Multi Choice, Freitext, Skala
- Zielgruppe: alle Lehrkräfte / Abteilung / manuelle Liste (Mails aus Stammdaten)
- Frist + Reminder
- Live-Ergebnis (nur für Ersteller), Export CSV
- Vorlagen: „Klassenwunsch nächstes Jahr“, „SCHILF-Themen“, „Stundenplan-Randbedingungen“

### Abgrenzung zum Einsammler

Gleiche Engine möglich (`Zielgruppe = lehrkraefte`), eigene Kachel „Kollegium“. Authentifiziert über MSAL → höhere Vertrauenswürdigkeit, keine Magic-Links.

### Lizenz / Preis

Paket **Kommunikation** oder **Eltern & Ablauf** · Wertbeitrag **90–130 €**/Schuljahr (im Bundle).

---

## 6. Produktsteckbrief: Ausgabenverwaltung (Schlüssel, Karten, Parkplatz)

**Umgesetzt (P5):** `tools/ausgaben.html`, SharePoint-Paket, Vorlagen + **freie Typen** (Typ als Text in SP). Details: [`ausgaben.md`](./ausgaben.md).

### Kurz

Einheitliches Tool: **Was wurde an wen ausgegeben?** Schlüssel, Raumkarten, Schrankenkarten, Parkplatz-/Zufahrtsberechtigungen – gleiche Logik, unterschiedlicher `GegenstandTyp`.

### Nutzen

Ersetzt Excel „Wer hat Schlüssel 12b?“. Berufsschulen: Parkplatz; AHS/HTL: Raum-/Schrankenkarten; alle: Generalschlüssel, Serverraum, Beamer-Koffer (optional später).

### Personas

Admin / Hausverwaltung / Sekretariat vergibt und nimmt zurück; optional Lehrkraft „meine Ausgaben“ (z. B. Klassenschlüssel).

### Funktionsumfang MVP

**Stammliste `AusgabenGegenstaende`** (Inventar):

| Feld | Beispiel |
|------|----------|
| Typ | `schluessel`, `raumkarte`, `schrankenkarte`, `parkplatz`, `sonstiges` |
| Kennung | Nr. / Kennzeichen / Karten-ID |
| Ort/Zone | Gebäude, Tor, Schrankreihe |
| Aktiv | ja/nein |

**Bewegungsliste `AusgabenVorgaenge`**:

| Feld | Bedeutung |
|------|-----------|
| GegenstandId / Kennung | Referenz |
| EmpfaengerTyp | `lehrer` \| `schueler` \| `extern` |
| EmpfaengerName / Email / KlasseCode | wem |
| AusgegebenAm / ZurueckAm | Zeitraum |
| Status | `ausgegeben` \| `retour` \| `verloren` \| `ersetzt` |
| KautionHinweis | nur Textflag, **kein** Zahlungsfluss |
| UnterschriftEmpfang | optional Ja/Bild |
| Notiz | Freitext |

UI:

- Inventar pflegen (Import CSV)
- „Ausgeben“-Dialog (Person + Gegenstand + Datum)
- „Zurücknehmen“ mit einem Klick
- Filter: offen ausgegeben, Typ, Person, überfällig (wenn `FaelligAm`)
- Ausdruck/PDF Übergabeprotokoll (einfach)

### Nicht im Scope

Zahlung/Kaution einziehen, physische Schranken-API, Kennzeichen-Erkennung, MDM.

### Erweiterung später (gleiche Listen)

Laptop-/Tablet-Ausgabe, Beamer-Koffer, Sportgeräte – nur neue `Typ`-Choices.

### Lizenz / Preis

Eigenes kleines Modul **„Ausgaben / Zugang“** unter Intranet oder Schul-Apps · **120–160 €**/Schuljahr (hoher Alltagsnutzen, geringer Support).

---

## 7. Paketierung & Roadmap

### 7.1 Vorgeschlagenes Lizenz-Paket (neu)

| Paket-ID | Label | Inhalt |
|----------|-------|--------|
| `ablauf` | **Eltern & Ablauf** | Einsammler, Elternbrief, Krankmeldung, Kollegiumsumfragen |
| `ausgaben` | **Ausgaben & Zugang** | Schlüssel/Karten/Parkplatz-Inventar |

Alternativ: alles unter erweitertem `schulapps` – dann Preiserhöhung Schul-Apps statt neuer SKUs.

**Preisbundle-Empfehlung:** Eltern & Ablauf **199 €**, Ausgaben **149 €**, zusammen **299 €**/Schuljahr.

### 7.2 Lieferreihenfolge

| Phase | Lieferobjekt | Warum zuerst |
|-------|--------------|--------------|
| **P0** | Plattform-Konventionen (dieses Doc) + Pakete `ablauf`/`ausgaben` im Catalog | **erledigt** |
| **P1** | Einsammler intern (ohne Public) | **erledigt** – `tools/einsammler.html` |
| **P2** | Public Magic-Link + Lesebestätigung-Modus | **erledigt** – `einsammler-antwort.html` |
| **P3** | Krankmeldung (Public + Inbox) | **erledigt** – `krankmeldung.html` / `krankmeldung-formular.html` |
| **P4** | Kollegiumsumfragen (MSAL-only) | **erledigt** – SharePoint SoR, Vorlagen, Auswertung |
| **P5** | Ausgabenverwaltung | **erledigt** – flexible Typen, Inventar + Vorgänge, SP-Paket |

### 7.3 Technische Wiederverwendung (bewusst bauen)

Beim Einsammler gleich mitdenken:

- `createCampaign({ audience, mode, questions })`
- `listRecipientsFromKlassen(klasseCodes)`
- `publicRespond({ token, payload })`
- `progressStats(campaignId)`

Krankmeldung und Kollegium hängen sich daran oder kopieren das Pattern 1:1.  
Ausgaben = separates Mini-Framework `issueItem` / `returnItem` / `openIssues`.

---

## 8. Abgrenzung zum Markt (Verkaufssatz)

| Konkurrent | Wir |
|------------|-----|
| WebUntis Student / Sprechtag | Die haben Plan & Kurswahl; wir ergänzen **Einwilligung, Lesen, Krank, Ausgaben** im M365-Tenant |
| edu.FLOW / Klasso | Ähnlicher Schmerz; wir: **Daten in SharePoint der Schule**, einheitlich mit übrigen Schultools |
| Schulmanager-Module | Vergleichbare Einzelpreise; wir als Bundle + bestehende Lizenzkunden |

Verkaufssatz (Entwurf):  
*„Untis plant den Unterricht – MS365-Schultools erledigt die Zettel: Einwilligungen, Lesebestätigungen, Krankmeldungen und wer welchen Schlüssel hat.“*

---

## 9. Offene Entscheidungen (kurz)

1. Neues Lizenzpaket `ablauf`/`ausgaben` vs. Aufpreis auf `schulapps`?  
2. Public-Forms: nur Graph-delegated Setup + später Azure-Function (wie Stammdaten-Portal) oder zuerst „manuell abhaken“?  
3. Elternbrief eigene Kachel oder nur Modus im Einsammler? → **eigene Kachel/Tool**  

4. Ausgaben: Kaution nur als Checkbox/Notiz – bestätigt **ohne** Payment?

---

## 10. Nächster konkreter Umsetzungsschritt

Wenn Freigabe: **P0+P1** – `ablauf` im Feature-Catalog, Schema `einsammler-schema.js`, SharePoint-Listen-Setup, interne UI „Einsammler anlegen + Antworten manuell setzen“. Public-Link folgt in P2.
