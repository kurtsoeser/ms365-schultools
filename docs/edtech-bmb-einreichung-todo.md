# ToDo: EdTech-Einreichung beim BMB (Bildungsportal)

**Ziel:** Stammdatenabruf über das Bildungsportal freigeben lassen und technisch umsetzen.  
**Nicht:** Direktanfrage bei Sokrates / bit media.  
**Begleitdokument:** [`edtech-bildungsportal-stammdatenanbindung.md`](./edtech-bildungsportal-stammdatenanbindung.md)

Status-Legende: `[ ]` offen · `[~]` in Arbeit · `[x]` erledigt

---

## Phase A – Vorbereitung (vor dem Support-Ticket)

### A1 Organisation & Rechtsträger
- [ ] Rechtsträger festlegen, der den EdTech-Vertrag unterzeichnet (Firma / Verein / …)
- [ ] Bevollmächtigte Person mit **ID-Austria** bestimmen (Onboarding + Unterzeichnung)
- [ ] Offizielle Kontaktdaten für BMB bereithalten (Firma, Adresse, Telefon, E-Mail, optional Logo)
- [ ] Interne Datenschutz-/Rechtsfreigabe: Zweck Stammdatenabruf + AVV ist gewollt und vertretbar
- [ ] Klären: Wer ist Auftragsverarbeiter gegenüber der Schule, wer Verantwortlicher (Formulierung für AVV)

### A2 Anwendung inhaltlich scharf schneiden
- [ ] Finalen **Anwendungsnamen** festlegen (Anzeige im Portal)
- [ ] Kurzbeschreibung ≤ 512 Zeichen formulieren (siehe Begleitdokument §3)
- [ ] Support-Kontakte der App festlegen (E-Mail, optional Telefon, Support-URL)
- [ ] Anwendungs-Logo bereitstellen
- [ ] Zweckbeschreibung für **Anhang B** finalisieren (nur Stammdaten für MS-365-Strukturpflege)
- [ ] Feldliste „minimal nötig“ aufschreiben (Datensparsamkeit): z. B. Name, Vorname, Rolle, Klasse, stabile ID, schulische E-Mail, Ein-/Austritt
- [ ] Explizit **nicht** beantragen: `manageuserdata_v1`, Mitteilungen, Zustellung, Content-Repos, Eltern/`lgn` (vorerst)
- [ ] Primär-APIs in Kurzbeschreibung nennen: `readorgdata` + `readuserdata_v3`

### A3 Technische Vorentscheidungen
- [ ] Entscheiden: API-Aufrufe über **kleines Backend** (empfohlen wegen Public Key / IP) oder anderer sicherer Weg
- [ ] Server-/Ausgangs-IPs für **IP-Einschränkung** notieren (sobald bekannt)
- [ ] Schlüsselpaar erzeugen und **Public Key** für BIP-Anwendung bereithalten
- [ ] Private Key sicher verwahren (nicht ins Git-Repo)
- [ ] Swagger grob sichten: https://www.bildung.gv.at/swagger
- [ ] Q-Umgebung als erstes Integrationsziel festlegen (keine Echtdaten vor Vertrag/AVV)

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

### C1 Partnerschaftsvertrag
- [ ] Vertragspartnerdaten eintragen
- [ ] Anwendungsbezug klar auf MS365-Schul-Tools beziehen

### C2 Anhang A – AVV
- [ ] AVV vollständig ausfüllen
- [ ] Verarbeitungsgegenstand: Stammdaten/Nutzerdaten aus Bildungsportal für Schul-IT-Provisioning
- [ ] Interne Rechtsprüfung / Unterschriftsfähigkeit sicherstellen

### C3 Anhang B – Schnittstellen (Kern der Einreichung)
- [ ] Anwendung beschreiben
- [ ] Zwecke der Datenverarbeitung beschreiben (Import/Abgleich Stammdaten → MS-365-Strukturen)
- [ ] **Primär beantragen:** `readuserdata_v3` (Personen) **und** `readorgdata` (Schule/Klassen)
- [ ] Nutzertypen beantragen: mindestens **`std` + `tch`** (Schüler:innen + Lehrkräfte)
- [ ] **Sekundär optional:** `searchuserdata_v3` (Punktabfrage, z. B. über `sokratesid`)
- [ ] **Nicht beantragen:** `manageuserdata_v1` (Schreiben für Schulanmeldung – anderer Use-Case)
- [ ] In Anhang B explizit technische Namen nennen (Swagger/OpenAPI), damit Freigabe eindeutig ist
- [ ] Datensparsamkeit und Zweckbindung explizit formulieren
- [ ] Widgets (Anhang C): vorerst „keine“ oder „später“ – klarstellen
- [ ] Anhang D (eigene Schnittstellen/Moodle): vorerst leer / nicht beantragen
- [ ] Anhang E (SSO-Token-Anreicherung): nur ausfüllen, wenn SSO von Anfang an geplant – sonst „noch nicht“

### C4 Interne Freigabe vor Absenden
- [ ] Alle Felder auf Vollständigkeit prüfen
- [ ] Keine übertriebenen Datenwünsche (erhöht Ablehnungs-/Verzögerungsrisiko)
- [ ] PDF/DOCX-Paket schnüren: Vertrag + A + B–E + ggf. Kurzcover (Begleitdokument §8)

---

## Phase D – Einreichung beim BMB

- [ ] Ticket/Anfrage über https://bip.gv.at/support erstellen
- [ ] Betreff klar: „EdTech-Partnerschaft – Antrag Stammdatenabruf …“
- [ ] Textbaustein aus Begleitdokument §8 verwenden und personalisieren
- [ ] Alle Anhänge hochladen / beifügen
- [ ] Bestätigung/Ticketnummer sichern
- [ ] Optional parallel: support@bildung.gv.at nur nutzen, wenn Support das so lenkt (Haupteinreichung bleibt bip.gv.at/support)

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

1. ID-Austria der bevollmächtigten Person  
2. Firmendaten + Kurzbeschreibung der App  
3. Entscheidung: Backend ja/nein + wer den Key hält  
4. Ausgefüllte Vertragsdokumente (A + B–E)  
5. Support-Ticket auf bip.gv.at/support  

Alles Weitere (Swagger-Details, Feldmapping, Knopf-UI) kann parallel oder nach Q-Zugang laufen.

---

## Merksatz für alle Gespräche mit dem Ministerium

> Wir beantragen **keine Sokrates-Direktintegration**, sondern die Bildungsportal-Standardschnittstellen **`readorgdata`** (Schul-/Klassenstammdaten) und **`readuserdata_v3`** (Personen, Nutzertypen `std`/`tch`), optional `searchuserdata_v3` – damit Schulen die bereits im Datenverbund liegenden Daten in MS365-Schul-Tools abgleichen können. **`manageuserdata`** beantragen wir nicht.
