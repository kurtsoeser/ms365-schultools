# ToDo: EdTech-Einreichung beim BMB (Bildungsportal)

**Ziel:** Stammdatenabruf über das Bildungsportal freigeben lassen und technisch umsetzen.  
**Nicht:** Direktanfrage bei Sokrates / bit media.  
**Begleitdokument:** [`edtech-bildungsportal-stammdatenanbindung.md`](./edtech-bildungsportal-stammdatenanbindung.md)

Status-Legende: `[ ]` offen · `[~]` in Arbeit · `[x]` erledigt

---

## Phase A – Vorbereitung (vor dem Support-Ticket)

### A1 Organisation & Rechtsträger
- [x] Rechtsträger festlegen, der den EdTech-Vertrag unterzeichnet (Firma / Verein / …)
  - **Kurt Söser**, Einzelunternehmen / Marke **#kurtrocks.edu.innovation** (GISA 29320646)
  - Adresse: Berggasse 75, 4400 Steyr
- [x] Bevollmächtigte Person mit **ID-Austria** bestimmen (Onboarding + Unterzeichnung)
  - **Kurt Söser**
- [x] Offizielle Kontaktdaten für BMB bereithalten (Firma, Adresse, Telefon, E-Mail, optional Logo)
  - Tel. **+43 670 1951157** · Mail **kontakt@kurtrocks.com** · Logo: `public/assets/ms365-schulverwaltung-logo.png`
- [x] Interne Datenschutz-/Rechtsfreigabe: Zweck Stammdatenabruf + AVV ist gewollt und vertretbar
- [x] Klären: Wer ist Auftragsverarbeiter gegenüber der Schule, wer Verantwortlicher (Formulierung für AVV)
  - **Schule** = Verantwortliche; **Kurt Söser / MS365-Schul-Tools** = Auftragsverarbeiter, soweit BIP-Abruf über betriebenes Backend; bei rein lokalem Browser-Abruf bleibt die Verarbeitung bei der Schule (Softwarebereitstellung).

### A2 Anwendung inhaltlich scharf schneiden
- [x] Finalen **Anwendungsnamen** festlegen (Anzeige im Portal) → **MS365-Schul-Tools**
- [x] Kurzbeschreibung ≤ 512 Zeichen formulieren (siehe Begleitdokument §3)
  - Browserbasierte Schul-IT-Werkzeuge zur Pflege von Microsoft-365-Strukturen (Gruppen, Teams, Jahrgänge, Kursteams, Konten-/Mitgliedschaftsabgleiche). Stammdatenabgleich über die Bildungsportal-Standardschnittstellen readorgdata und readuserdata_v3 (Schüler:innen und Lehrkräfte). Verarbeitung lokal im Browser bzw. über Microsoft Graph im Auftrag der Schule; kein zentraler App-Stammdatenspeicher.
- [x] Support-Kontakte der App festlegen (E-Mail, optional Telefon, Support-URL)
  - Mail **kontakt@kurtrocks.com** · Tel. **+43 670 1951157** · URL **https://app.ms365.schule/hilfe.html**
- [x] Anwendungs-Logo bereitstellen → `public/assets/ms365-schulverwaltung-logo.png`
- [x] Zweckbeschreibung für **Anhang B** finalisieren (nur Stammdaten für MS-365-Strukturpflege)
- [x] Feldliste „minimal nötig“ aufschreiben (Datensparsamkeit)
  - Nachname, Vorname, Rolle (`std`/`tch`), Klasse(n), stabile ID (`sokratesids`/`bpkbf`), schulische E-Mail (falls freigegeben), Status / Ein-/Austritt
- [x] Explizit **nicht** beantragen: `manageuserdata_v1`, Mitteilungen, Zustellung, Content-Repos, Eltern/`lgn` (vorerst)
- [x] Primär-APIs in Kurzbeschreibung nennen: `readorgdata` + `readuserdata_v3`

### A3 Technische Vorentscheidungen
- [x] Entscheiden: API-Aufrufe über **kleines Backend** (empfohlen wegen Public Key / IP) oder anderer sicherer Weg
  - **Entscheidung:** BIP wird **nicht** aus dem Browser aufgerufen.
  - **Architektur:** kleines **BIP-Proxy-Backend** (`backend/bip-api`, analog zu `license-api` / `kursteams-api`): hält Private Key + BIP-Credentials, ruft `readorgdata` / `readuserdata_v3` serverseitig auf, liefert Abgleichsdaten an die App.
  - **Hosting (feste Ausgangs-IP):** **günstiger EU-VPS mit fester IPv4** – Erstwahl **Hetzner Cloud** (ca. 4–6 €/Monat). Azure mit Static Outbound (VM ~15–25 €, Premium+NAT ~200 €) bewusst verworfen (zu teuer für den Use-Case). World4you-Webhosting ungeeignet (keine stabile dedizierte Outbound-IP für BIP).
  - **Key-Halter:** Kurt Söser / Server – Private Key nur auf dem VPS (Secret/Dateirechte), nie im Frontend, nie im Git.
- [~] Server-/Ausgangs-IPs für **IP-Einschränkung** notieren (sobald bekannt)
  - **Jetzt (Formular speichern):** Arbeitsplatz-IP `86.56.206.76` (Übergang, dynamisch)
  - **Vor Q-Integration:** feste VPS-IPv4 eintragen und Arbeitsplatz-IP entfernen (sobald VPS angelegt)
  - Platzhalter in BIP-UI bis dahin ok; vor Echttests zwingend ersetzen
- [x] Schlüsselpaar erzeugen und **Public Key** für BIP-Anwendung bereithalten
  - Ablage außerhalb Git: `%USERPROFILE%\.ms365schule-secrets\bip\`
  - Dateien: `bip-schnittstellen-public.pem` / `bip-schnittstellen-private.pem` (RSA 4096)
- [x] Private Key sicher verwahren (nicht ins Git-Repo)
- [x] Swagger grob sichten: https://www.bildung.gv.at/swagger
  - Bestätigt laut OpenAPI: Basic Auth; P1 = `readorgdata` + `readuserdata_v3`; optional `searchuserdata_v3`
- [x] Q-Umgebung als erstes Integrationsziel festlegen (keine Echtdaten vor Vertrag/AVV)
  - Reihenfolge: Vertrag/AVV/Anhang B → Freischaltung → **Q** → erst danach I/P

#### A3 Nacharbeit (Hosting, **verschoben** – nicht blockierend für Formular/Einreichung)
- [ ] ~~jetzt~~ später: Hetzner-Cloud-Konto / kleinster sinnvoller VPS anlegen (feste IPv4 notieren)
- [ ] BIP-Formular: IP-Einschränkung auf VPS-IP umstellen
- [ ] Private Key + BIP-Credentials auf VPS ablegen
- [ ] Später Phase F: `backend/bip-api` implementieren und deployen
- Status 2026-09-28: Hosting bewusst **zurückgestellt**; Einreichung läuft mit Übergangs-IP `86.56.206.76` weiter.

---

## Phase B – Onboarding im Bildungsportal

### B1 Partner & Anwendung anlegen
- [ ] Mit ID-Austria öffnen: https://bip.gv.at/edtech/onboarding
- [ ] EdTech-**Partner** anlegen (Unternehmens-/Vereinsdaten, nicht Privatadresse)
- [ ] **Anwendung** anlegen: Name, Beschreibung, Support, Logo
- [ ] Public Key (Schnittstellen) hinterlegen
- [ ] IP-Einschränkung setzen (sobald IPs feststehen)
- [ ] Weitere Berechtigte einladen (Menü „Rechte“), falls nötig
- [ ] Prüfen: Organisation „edTech Partner“ am Dashboard sichtbar; Link zur Tech-Doku / Infobox

### B2 Dokumentation laden
- [ ] FAQ EdTech-Vereinbarung lesen: https://www.bildung.gv.at/filter/faq/page.php?p=168
- [ ] FAQ Schnittstellen lesen: https://www.bildung.gv.at/filter/faq/page.php?p=167
- [ ] Partnerschaftsvertrag (docx/odt) herunterladen
- [ ] Anhang A AVV herunterladen
- [ ] Anhänge B–E (Schnittstellen, Widgets, SSO) herunterladen
- [ ] Infobox/Tech-Doku nach Login öffnen: https://bip.gv.at/infobox

---

## Phase C – Vertragsunterlagen ausfüllen

**Entwürfe (2026-09-28):** `docs/edtech-vertragsentwurf/ausgefuellt-*.docx` (+ Kopie in Downloads)  
**Ausfüllhilfe:** `docs/edtech-vertragsentwurf/AUSFUELLHILFE.md`

### C1 Partnerschaftsvertrag
- [x] Vertragspartnerdaten eintragen → Kurt Söser / GISA / Steyr
- [x] Anwendungsbezug klar auf MS365-Schul-Tools beziehen *(über Anhänge)*
- [ ] Datum eintragen + Unterschrift (nach Prüfung / vor finaler Einreichung)

### C2 Anhang A – AVV
- [x] Partnerkopf ausfüllen
- [x] Verarbeitungsgegenstand: Stammdaten/Nutzerdaten aus Bildungsportal für Schul-IT-Provisioning *(via Anhang B)*
- [x] Anlage 1 TOM: relevante Checkboxen + Sonstige-Texte
- [ ] Interne Rechtsprüfung / Unterschriftsfähigkeit sicherstellen
- [ ] Datum + Unterschrift

### C3 Anhang B – Schnittstellen (Kern der Einreichung)
- [x] Anwendung beschreiben
- [x] Zwecke der Datenverarbeitung beschreiben
- [x] **Primär beantragen:** `readuserdata_v3` + `readorgdata` → **an aktivierten Schulen**
- [x] Nutzertypen: **`std` + `tch`**
- [x] **Sekundär:** `searchuserdata_v3` → an aktivierten Schulen
- [x] **Nicht beantragen:** App-Berechtigungen, Benachrichtigungen, Amtssignatur, `manageuserdata`
- [x] Technische Endpoint-Namen in Umfangstexten
- [x] Datensparsamkeit formuliert
- [x] Widgets (Anhang C): keine / später
- [x] Anhang D: keine
- [x] Anhang E (SSO): nicht beantragt

### C4 Interne Freigabe vor Absenden
- [x] Paket eingereicht (2026-09-28) – visuelle Word-Prüfung vor Absenden vorausgesetzt
- [x] Keine übertriebenen Datenwünsche
- [x] DOCX-Paket: Vertrag + A + B–E
- [x] Support-Ticket-Text aus Begleitdokument verwendet

---

## Phase D – Einreichung beim BMB

- [x] Ticket/Anfrage über https://bip.gv.at/support erstellen (**2026-09-28**)
- [x] Betreff: EdTech-Partnerschaft – Antrag Stammdatenabruf (readorgdata + readuserdata_v3) für MS365-Schul-Tools
- [x] Textbaustein personalisiert und gesendet
- [x] Anhänge: ausgefuellt Partnerschaftsvertrag + Anhang A + Anhang B_E
- [x] Bestätigung/Ticketnummer sichern → **#12337** (2026-09-28)
- [x] Haupteinreichung über bip.gv.at/support

---

## Phase E – Nachkontakt & Freischaltung

- [ ] Zugewiesene Ansprechpersonen des BMB notieren
- [ ] Rückfragen zu Zweck/Datensparsamkeit beantworten
- [ ] Einvernehmlich freigegebene Schnittstellenliste dokumentieren (was genau freigegeben wurde)
- [ ] Partnerschaftsvertrag beiderseits unterzeichnen lassen
- [ ] AVV unterzeichnen lassen
- [ ] Im Partner-Dashboard prüfen: Schnittstellen für die Anwendung freigeschaltet
- [ ] Zugang Q-Umgebung / Credentials / Client-IDs laut Infobox/Swagger notieren
- [ ] Erst danach I/P anfragen bzw. freischalten lassen

---

## Phase F – Technische Umsetzung in der App (nach Freigabe / parallel in Q)

### F1 Integration
- [ ] OpenAPI laden: https://www.bildung.gv.at/local/eduportal/iface/openapi.php
- [ ] Auth laut Partner-Doku / OpenAPI (Basic Auth + freigeschaltete App; Public Key / IP laut Onboarding)
- [ ] Client für **`readorgdata`**: Klassen/Abteilungen/Schulname vor Personenimport
- [ ] Client für **`readuserdata_v3`**: `orgids`, `usertypes=std,tch`, `timechanged`, Pagination
- [ ] Optional Client für **`searchuserdata_v3`** (Matching über `sokratesid` / bPK)
- [ ] Mapping: Org-Klassenliste + Personen (`sokratesids`/`bpkbf`, `orgs[].classes`) → App-Stammdaten
- [ ] Abgleich-Logik: neu / geändert / `deleted`/`suspended`; letzten Sync-Timestamp speichern
- [ ] UI: Knopf „Stammdaten mit Bildungsportal abgleichen“ (+ Vorschau vor dem Übernehmen)
- [ ] Fehler-/Rechtefälle: Schule nicht freigeschaltet, leere Daten, Teilrechte
- [ ] Logging/Audit schulseitig: wer hat wann synchronisiert (ohne unnötige personenbezogene Logs)

### F2 Pilot
- [ ] Pilotschule festlegen (gute Stammdatenqualität im Register)
- [ ] Testplan: Erstimport, Delta-Abgleich, Klassenwechsel, Austritt
- [ ] Datenschutzhinweis in `hilfe.html` / Datenschutzerklärung ergänzen (BIP-Abruf)
- [ ] Betriebsanleitung für Schul-IT: wer darf den Knopf nutzen

### F3 Parallel: Unterricht / Kursteams über WebUntis (nicht BIP)
- [ ] Klar trennen: BIP = Personen/Schule/Klassen; **Unterricht/Kursteams = WebUntis oder CSV**
- [ ] **Default:** bestehender Kursteams-Import (CSV/Excel) beibehalten
- [ ] Optional: Untis-Partner anfragen – Formular https://www.untis.at/integrationen/kontaktformular-integrationspartner
- [ ] Doku: https://developer.untis.com/getting-started/get-platform-application/
- [ ] Entscheidung: BIP zuerst; Untis-API nur wenn Partnerzugang kommt – sonst CSV
- [ ] Nicht erwarten, dass BIP OpenAPI Unterricht/Stundenplan liefert (Widget ≠ Datendrehscheibe)

---

## Sofort-Packliste (was du diese Woche brauchst)

Zum Starten reichen:

1. ID-Austria der bevollmächtigten Person → **Kurt Söser**  
2. Firmendaten + Kurzbeschreibung der App → **erledigt (A1/A2)**  
3. Entscheidung: Backend ja/nein + wer den Key hält → **BIP-Proxy auf VPS (Hetzner), Key auf Server (A3)**  
4. Ausgefüllte Vertragsdokumente (A + B–E)  
5. Support-Ticket auf bip.gv.at/support  

Formular jetzt: Public Key aus `%USERPROFILE%\.ms365schule-secrets\bip\bip-schnittstellen-public.pem`, IP vorerst `86.56.206.76`, Hilfe-URL `https://app.ms365.schule/hilfe.html`.

Alles Weitere (VPS-IP final, Feldmapping, Knopf-UI) kann parallel oder nach Q-Zugang laufen.

---

## Merksatz für alle Gespräche mit dem Ministerium

> Wir beantragen **keine Sokrates-Direktintegration**, sondern die Bildungsportal-Standardschnittstellen **`readorgdata`** (Schul-/Klassenstammdaten) und **`readuserdata_v3`** (Personen, Nutzertypen `std`/`tch`), optional `searchuserdata_v3` – damit Schulen die bereits im Datenverbund liegenden Daten in MS365-Schul-Tools abgleichen können. **`manageuserdata`** beantragen wir nicht.
