# Stammdatenanbindung über Bildungsportal / EdTech Hub

**Zweck:** Entscheidungsgrundlage, API-Referenz und Einreichungsunterlage für die Anbindung der MS365-Schulverwaltung an die **Standardschnittstellen des Bildungsportals** (BMB) – **nicht** an Sokrates direkt.

**Stand:** September 2026 (Swagger/OpenAPI ausgewertet)  
**Zielbild:** Ein Knopf in der App, der Stammdaten (Schule/Klassen, Schüler:innen, Lehrkräfte) mit den führenden Schuldaten abgleicht bzw. importiert.

**Verwandt:** ToDo-Checkliste [`edtech-bmb-einreichung-todo.md`](./edtech-bmb-einreichung-todo.md)

---

## 1. Kurzfassung

| Frage | Antwort |
|--------|---------|
| Ministeriums-Initiative | **Bildungsportal** (`bildung.gv.at`) / **EdTech Hub** |
| Datenherkunft „aus Sokrates“ | Sokrates → **SV-REG / Datenverbund Schule** → Bildungsportal → Partner-API |
| Andocken bei | **BMB / BIP** als EdTech-Partner |
| Kern-APIs (P1) | `readorgdata` + `readuserdata_v3` |
| Optional (P2) | `searchuserdata_v3` |
| Nicht beantragen | `manageuserdata_v1` (Schreiben für Schulanmeldung) |
| Rechtlich nötig | Partnerschaftsvertrag + **AVV** + Freigabe in **Anhang B** |

---

## 2. Architektur

```
Sokrates                    Untis Desktop (Gruber & Petters)
  │                                │
  │ Personen                       │ Unterricht, Stundenplan, Fächer, …
  ▼                                ▼
SV-REG / Datenverbund          WebUntis
  │                                │
  │ Stammdaten                     ├─ Widget/SSO → Anzeige im BIP-Dashboard
  ▼                                └─ Platform API → Kursteams / Unterrichtsdaten für Partner
Bildungsportal (BIP)
  │
  ├─ readorgdata / readuserdata_v3  → Personen, Schule, Klassen-Namen
  └─ (kein Unterrichts-Endpoint in der Partner-OpenAPI)
                │
                ▼
         MS365-Schulverwaltung
         · Knopf Stammdaten (BIP)
         · Kursteams / Unterricht (WebUntis API oder CSV)
```

Politisch und vertraglich andocken wir am **Bildungsportal**, nicht bei Sokrates. Die Datenqualität kommt aus der Schulverwaltung (häufig Sokrates); Freigabe und Vertrag laufen über das BMB.

Bereits angebundene Systeme (BMB-Auskunft 2025 u. a.): Sokrates (Bund/Privat), UNTIS/WebUntis, PM-SAP, TeachersDirect, eduvidual, lms.at, ABA-Portal.

---

## 3. Zweckbeschreibung für Anhang B (Arbeitsstand)

**Anwendungsname:** MS365-Schulverwaltung / MS365-Schultools  

**Kurzbeschreibung (≤ 512 Zeichen, Entwurf):**  
Browserbasierte Schul-IT-Werkzeuge zur Pflege von Microsoft-365-Strukturen (Gruppen, Teams, Jahrgänge, Kursteams, Konten-/Mitgliedschaftsabgleiche) auf Basis lokal gehaltener Stammdaten. Stammdaten werden nicht an einen App-Server der Anwendung gesendet; Verarbeitung erfolgt lokal im Browser bzw. über Microsoft Graph im Auftrag der Schule.

**Beantragte Verarbeitungszwecke:**

1. Import bzw. Abgleich von **Schul-/Klassen- und Personenstammdaten**, um Doppelpflege zu vermeiden.
2. Unterstützung der Einrichtung und Pflege von **Microsoft-365-Gruppen/Teams** auf Basis aktueller schulischer Stammdaten.
3. **Aktualisierung** bei Namens-, Klassen- oder Statusänderungen (inkl. Austritt/`deleted`/`suspended`), soweit freigegeben.

**Datensparsamkeit:** Nur Identifikation und Strukturpflege – Name, Rolle, Klasse, stabile IDs (`bpkbf` / `sokratesids`), schulische E-Mail falls freigegeben, Status. Keine Fotos, keine Privatadressen, keine Erziehungsberechtigten (vorerst), keine Weitergabe außerhalb der schulischen MS-365-Umgebung und lokalen App-Nutzung.

---

## 4. API-Katalog (Entscheidungstabelle)

Quellen:

- Swagger-UI: https://www.bildung.gv.at/swagger/
- OpenAPI: https://www.bildung.gv.at/local/eduportal/iface/openapi.php
- Basis-URL: `https://www.bildung.gv.at`
- Auth laut OpenAPI: **Basic Auth** (zusätzlich Partner-Public-Key / IP laut Onboarding)

| Prio | Endpoint | Richtung | Rolle für uns |
|------|----------|----------|---------------|
| **P1** | `…/local_eduportal_iface_readorgdata` | Lesen | Schulstammdaten + **Klassenliste** |
| **P1** | `…/local_eduportal_iface_readuserdata_v3` | Lesen | **Personen** Bulk + Delta (`timechanged`) |
| **P2** | `…/local_eduportal_iface_searchuserdata_v3` | Lesen | Einzelperson (z. B. `sokratesid`) |
| **nein** | `…/local_eduportal_iface_manageuserdata_v1` | Schreiben | Schulanmeldung – anderer Use-Case |

Vollständige Pfade beginnen jeweils mit:  
`/local/eduportal/webservice/server.php/`

**Nutzertypen (`usertypes`):**

| Code | Bedeutung | Antrag |
|------|-----------|--------|
| `std` | Schüler:in | **ja** |
| `tch` | Lehrkraft | **ja** |
| `dir` | Direktion / Schulleitung | optional |
| `lgn` | Erziehungsberechtigte / Bezugsperson | vorerst **nein** |

---

## 5. Sync-Ablauf (Zielbild Knopf)

```
1) readorgdata(orgids = Schulkennzahl)
      → Schulname, classes[], departments[]
      → Klassen in App-Stammdaten anlegen/aktualisieren

2) readuserdata_v3(orgids, usertypes=std,tch, timechanged=…)
      → Personen + orgs[].roles / orgs[].classes
      → Abgleich über sokratesids bzw. bpkbf

3) optional searchuserdata_v3(sokratesid | bpkbf | …)
      → manuelles Nachziehen / Konfliktlösung / Debug
```

**Delta:** Speichere den höchsten verarbeiteten `timechanged` (bzw. Sync-Zeitpunkt). Folgeläufe mit diesem Wert, damit nur Änderungen kommen. Pagination über `limitnum` + `cursor` / `meta.next_cursor` / `meta.has_more`.

---

## 6. Schnittstelle im Detail: `readorgdata` (P1)

**Summary:** Stammdaten über Schulen abrufen  

**Request:**

```json
{ "orgids": "900001" }
```

- `orgids` leer → alle Schulen, auf die die Anwendung Zugriff hat  
- Mehrere Schulkennzahlen komma-getrennt möglich  

**Response (Maximum laut Spec – Freigabe kann kürzen):**

| Feld | Nutzen für uns |
|------|----------------|
| `orgid` | Schulkennzahl |
| `name` / `officialname` / `shortname` | Anzeigenamen |
| `email` / `phone` / Adresse | optionale Schulstammdaten |
| **`classes[]`** | Klassenliste des Schuljahres |
| `classcount` | Plausibilität |
| **`departments[]`** | Abteilung + zugehörige Klassen |
| `orgtype`, `eduregion`, `educluster` | Kontext / Filter |
| `genuine` / `operational` | Demo vs. echt / aktiv |
| `timechanged` | Änderungsstand Org |

**Warum P1:** Klassenstruktur vor Personenimport – passt zu Tenant-Stammdaten, Jahrgängen, Kursteams.

---

## 7. Schnittstelle im Detail: `readuserdata_v3` (P1 – Kern)

**Summary:** Benutzerdaten lesen (Bulk)  

**Request:**

```json
{
  "orgids": "900001",
  "usertypes": "std,tch",
  "timechanged": 0,
  "limitnum": 50000,
  "cursor": ""
}
```

| Parameter | Bedeutung |
|-----------|-----------|
| `orgids` | Schulkennzahlen; leer = alle zugänglichen Orgs |
| `usertypes` | `std`, `tch`, `dir`, `lgn` |
| `timechanged` | nur seit Timestamp geänderte Nutzer; `0` = alle |
| `limitnum` | max. Anzahl (Pagination) |
| `cursor` | Cursor aus vorheriger Antwort |

**Response – für uns relevante Felder (Minimalset beantragen):**

| Feld | Nutzen |
|------|--------|
| `bpkbf` | stabiler Identifikator (BF:…) |
| `sokratesids[]` | Abgleich mit Sokrates-/Altbeständen |
| `firstname`, `lastname` (ggf. `middlename`, Titel) | Anzeigename |
| `emails[]` | MS-365-Matching (schulisch / preferred) |
| `orgs[].orgid` | Schulzuordnung |
| `orgs[].roles[]` | Rolle an der Schule |
| `orgs[].classes[].name` | Klassenzuordnung |
| `orgs[].departments[].name` | Abteilung |
| `deleted`, `suspended` | Austritt / Sperre → Abgleich |
| `timechanged` | Delta-Steuerung |
| `idpusername` | ggf. lokaler Nutzername |
| `idp_subjectid_v1` | BIP-Identität |
| `meta.has_more` / `meta.next_cursor` | Pagination |

**Response – vorerst nicht beantragen / nicht speichern (Datensparsamkeit):**

- `photo_*`, `addresses`, `phonenumbers` (außer klarer Bedarf)
- `children`, `relatives` / Erziehungsberechtigte
- `dateofbirth`, `gender` nur wenn wirklich für MS-365 nötig (meist nein)

**Hinweis Spec:** Die Doku zeigt das Maximum; Attribute und Usertypes können pro Anwendung reduziert werden.

---

## 8. Schnittstelle im Detail: `searchuserdata_v3` (P2)

**Summary:** Stammdaten über Personen suchen  

**Kein Ersatz für den Schul-Vollabzug.** Mindestens **ein** Suchkriterium:

- `bpkbf` **oder**
- `idp_subjectid_v1` **oder**
- `sapid` **oder**
- `sokratesid` **oder**
- `dateofbirth` **und** `zip`

Optional: `name` (teilweise, case-/akzent-insensitive), `usertypes`, `documentid`.

**Sinnvolle Nebenrollen:** Matching bei bekannter Sokrates-ID, manuelle Suche in der UI, Debug in Q.

Antwortstruktur weitgehend wie `readuserdata_v3` (ohne Bulk-Meta).

---

## 9. Nicht beantragen: `manageuserdata_v1`

**Summary:** Nutzerdaten verwalten (Schreiben)  

Use-Case laut Spec: **Schulanmeldung** – Partner schreiben Attribute (`lgn`, `communication`, `address`, Rollen `std_preregistered` / `std_admitted`) **ins** Portal, typischerweise nach `searchuserdata`.

Die reguläre Schülerrolle kommt aus dem Schülerverwaltungssystem und ist **nicht** Teil dieser Schnittstelle.

→ Anderer Zweck, anderes Risiko, andere Freigabe. Für MS365-Import **nicht** beantragen.

---

## 10. Mapping-Skizze → App-Stammdaten

| BIP | App (Arbeitsrichtung) |
|-----|------------------------|
| `readorgdata.classes[]` | Klassenstammdaten / Klassenliste |
| `readorgdata.name` / `orgid` | Schule / Schulkennzahl in Einstellungen |
| `readuserdata` Person + `sokratesids` / `bpkbf` | stabile Person-ID im lokalen Bestand |
| `firstname` + `lastname` | Anzeigename |
| `emails[]` (preferred/official) | Matching Entra/Graph-Benutzer |
| `orgs[].classes` | Klassenzuordnung Schüler:in |
| `usertypes` / `orgs[].roles` | Lehrkraft vs. Schüler:in |
| `deleted` / `suspended` | Austritt / deaktivieren in Abgleich |

Konkrete Feldnamen der App (`app-data` / Tenant) werden bei der Umsetzung in Phase F verfeinert.

---

## 11. Offizielle Einstiege & Vertragsweg

| Ressource | URL |
|-----------|-----|
| EdTech-Onboarding | https://bip.gv.at/edtech/onboarding |
| Support / Einreichung | https://bip.gv.at/support |
| FAQ Schnittstellen | https://www.bildung.gv.at/filter/faq/page.php?p=167 |
| FAQ EdTech-Vereinbarung | https://www.bildung.gv.at/filter/faq/page.php?p=168 |
| Swagger | https://www.bildung.gv.at/swagger/ |
| OpenAPI | https://www.bildung.gv.at/local/eduportal/iface/openapi.php |
| Infobox (nach Login) | https://bip.gv.at/infobox |
| Onboarding-PDF | https://oead.at/fileadmin/Dokumente/oead.at/Bildung_Digital/Marktplatz_Lernapps/Dateien/Regelbetrieb_2026-27/2026-02-23_BMB_Handreichung_edTech-Onboarding.pdf |
| BMB-Projektseite | https://www.bmb.gv.at/Themen/schule/zrp/dibi/bip.html |

**Umgebungen:** Q (Testdaten, oft vor Vertrag) → I/P erst nach unterzeichneter Vereinbarung + freigeschalteten Schnittstellen.

**Onboarding-Technik:** Public Key an der Anwendung, optional IP-Whitelist; SSO später möglich.

**Einreichungsschritte:** Partner anlegen → Anwendung → Vertrag + AVV + Anhänge B–E → `bip.gv.at/support` → Freischaltung → Q-Integration.

---

## 12. Entwurf Support-Nachricht

> Betreff: EdTech-Partnerschaft – Antrag Stammdatenabruf (readorgdata + readuserdata_v3) für MS365-Schulverwaltung  
>  
> Sehr geehrte Damen und Herren,  
>  
> wir möchten die Anwendung **MS365-Schulverwaltung** als EdTech-Partneranwendung an das Bildungsportal anbinden. Beantragt werden die Standardschnittstellen  
> **`local_eduportal_iface_readorgdata`** und **`local_eduportal_iface_readuserdata_v3`**  
> (optional ergänzend **`local_eduportal_iface_searchuserdata_v3`**),  
> Nutzertypen mindestens **`std`** und **`tch`**, schulbezogen über `orgids`.  
>  
> Zweck: Import/Abgleich von Schul-/Klassen- und Personenstammdaten zur Entlastung der Doppelpflege und zur Unterstützung der Microsoft-365-Strukturpflege an Schulen. Anbindung erfolgt über das Bildungsportal / den Datenverbund Schule – **nicht** über eine Direktintegration Sokrates.  
>  
> **Nicht beantragt:** `manageuserdata_v1` (Schreiben im Schulanmeldungsprozess).  
>  
> Beigefügt: Partnerschaftsvereinbarung, AVV (Anhang A), Anhänge B–E.  
> Wir bitten um Zuweisung der fachlichen Ansprechpersonen und Freigabe zur Integration in der Q-Umgebung.  
>  
> Mit freundlichen Grüßen  
> [Name, Organisation, Telefon, E-Mail]

---

## 13. Abgrenzung Sokrates / WebUntis (nicht Einreichungsweg)

| Weg | Für uns? | Bemerkung |
|-----|----------|-----------|
| Sokrates `…/ws/untis` | nein | Untis-spezifisch (`DataExchangeService`) |
| Sokrates CSV „Schüler für WebUntis“ | nur Fallback | manueller Export |
| WebUntis Platform APIs | **ja für Unterricht/Kursteams** | eigener Partnerweg bei Untis (Gruber & Petters) |
| **BIP `readorgdata` + `readuserdata_v3`** | **ja für Personen/Schule/Klassen** | Haupteinreichung BMB |

### 13.1 Warum ich meinen Unterricht im Bildungsportal sehe – und warum das *nicht* heißt, dass BIP die Unterrichtsdaten speichert

In der Praxis erscheint der Stundenplan/Unterricht im BIP-Dashboard. Laut BIP-FAQ ([WebUntis-Konfiguration](https://www.bildung.gv.at/filter/faq/page.php?lang=de&p=48&t=)) ist das die **WebUntis-Widget-Anbindung**:

> Die Schnittstelle zwischen WebUntis und dem Bildungsportal ermöglicht es, den **WebUntis-Stundenplan als Widget** im Bildungsportal-Dashboard anzuzeigen.

Das bedeutet fachlich:

| Schicht | Was passiert |
|---------|----------------|
| **Untis Desktop → WebUntis** | Stammdaten, **Unterricht**, Stundenplan, Vertretungen werden nach WebUntis übertragen (Produkt Gruber & Petters / Untis) |
| **WebUntis ↔ Bildungsportal** | vor allem **Identitätsabgleich** (Externe ID: Sokrates-ID bei Schüler:innen, SAP-Personalnummer bei Lehrkräften) + **Anzeige** des Plans als Widget / SSO |
| **BMB-Datenabgleich** | Personenstammdaten BIP ↔ WebUntis – **nicht** der Unterrichtskatalog |
| **BIP Partner-OpenAPI** | `readorgdata` / `readuserdata*` / `searchuserdata*` / … – **kein** Endpoint für Unterricht, Fächer, Stundenplan, Kurse |

**Folge für MS365-Kursteams:** Unterrichtsdefinitionen (Fach + Klasse/Schülergruppe + Lehrkraft = „lesson“) bleiben führend in **WebUntis**. Sie lassen sich über die öffentlichen BIP-Standardschnittstellen **nicht** auslesen. Dafür braucht es die **WebUntis Platform APIs** (z. B. Lessons / Student Management / Timetable) als **zweiten** Integrationspfad neben dem EdTech-Antrag beim BMB.

```
Personen / Schule / Klassen-Namen     →  BIP (readorgdata + readuserdata_v3)
Unterricht / Kurse / wer sitzt worin  →  WebUntis Platform API (Untis-Partner)
Widget „mein Stundenplan“ im BIP     →  nur Anzeige aus WebUntis, keine BIP-Datendrehscheibe Unterricht
```

Optional später: eigenes BIP-Widget der MS365-App (Anhang C) – das ersetzt nicht den Datenabruf aus WebUntis.

### 13.2 Unterricht / Kursteams: Praxisweg

| Option | Wann |
|--------|------|
| **CSV/Excel** (z. B. ExportLesson wie bisher) | **Standard jetzt** – ohne Untis-Partnerzugang sofort nutzbar |
| **WebUntis Platform API** | Nur nach Genehmigung als Integrationspartner |

**Untis-Partner beantragen (optional, parallel zu BIP):**

1. Formular: https://www.untis.at/integrationen/kontaktformular-integrationspartner  
2. Technik: https://developer.untis.com/getting-started/get-platform-application/  
3. Angeben: Produktname, Firma, Use-Case (Kursteams aus Lessons → MS 365), Länder/Schulanzahl, Integration vor allem **API** (Opt-in pro Schule)  
4. Kontakt: office@untis.at · +43 2266 62241 · Untis GmbH, Stockerau  

**Hinweis:** Zugang ist eher für Softwareprodukte mit mehreren Schulen gedacht. Ablehnung oder lange Wartezeit → bei CSV bleiben; BIP-Stammdaten davon unabhängig vorantreiben.

---

## 14. Offene Produktentscheidungen


1. Rechtsträger / Unterzeichner EdTech-Vertrag  
2. API-Aufrufe: kleines Backend (empfohlen: Keys, IP, Audit) vs. anderer sicherer Weg  
3. Pilotschule  
4. Ob bis zur Freigabe WebUntis/CSV als Übergang gebaut wird  

---

## 15. Verwandte Dokumente

- [`edtech-bmb-einreichung-todo.md`](./edtech-bmb-einreichung-todo.md) – Phasen A–F Checkliste  
- [`projektanalyse-und-werkzeugideen.md`](./projektanalyse-und-werkzeugideen.md) – App-Kontext  
