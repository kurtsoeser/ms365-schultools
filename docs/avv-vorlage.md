# AVV-Vorlage – MS365-Schul-Tools

**Status:** Vorlage für Schulen (Stand 21.11.2026). Kein Ersatz für individuelle Rechtsberatung.  
**Zweck:** Auftragsverarbeitung zwischen Schule (Verantwortliche) und Kurt Söser / MS365-Schul-Tools (Auftragsverarbeiter), soweit Betreiber-Dienste personenbezogene Daten verarbeiten.

---

## 1. Parteien

| Rolle | Organisation | Kontakt |
|---|---|---|
| Verantwortliche | [Schulname, Adresse] | [Datenschutz / IT] |
| Auftragsverarbeiter | Kurt Söser, Berggasse 75, 4400 Steyr, Österreich | kontakt@kurtrocks.com |

Marke / Produkt: **MS365-Schul-Tools** (ms365.schule / app.ms365.schule).

---

## 2. Gegenstand und Dauer

Der Auftragsverarbeiter stellt die Browser-App und schlanke Betriebsdienste bereit:

1. **Statisches Hosting** der App-Oberfläche (HTML/JS/CSS).
2. **Lizenz-API** zur Freischaltung (Tenant-ID, Module/Pakete, optionale Kontaktdaten, optional Site-URL-Zeiger).
3. **Optionale Portale/APIs** nur bei ausdrücklich freigeschaltetem Modul (z. B. Wahl, Einsammler-Remote) – Standard ist local-first / SharePoint der Schule.

Die Vereinbarung gilt für die Dauer der Nutzung bzw. bis zur Kündigung / Löschung der Betreiber-Daten.

---

## 3. Art der Daten und betroffene Personen

| Kategorie | Beispiele | Typisch betroffen |
|---|---|---|
| Betriebs-/Lizenzdaten | Tenant-ID, Freischaltungsstatus, Schulname, Kontakt-E-Mail, Rollenhinweis | IT, Direktion |
| Site-Zeiger | URL der schulischen Tool-/SharePoint-Site | IT |
| Optionale Portaldaten | nur bei freigeschaltetem Modul und dokumentiertem Zweck | je Modul |

**Nicht** Gegenstand dieser AVV als Produkt-Datenbank des Betreibers:

- Klassen-, Schüler-, Lehrerstammlisten
- Freistellungen, Schularbeiten, sRDP-Anmeldungen, Elternkanal-Belege in SharePoint der Schule
- lokale Browser-Stände (`localStorage`) der bedienenden Person

Diese Daten verarbeitet die Schule in Microsoft 365 bzw. lokal; der Betreiber speichert sie dort nicht als zentrale Produkt-Datenbank.

---

## 4. Weisungen und Zwecke

Der Auftragsverarbeiter verarbeitet Betreiber-Minimaldaten nur zur:

- Lizenzprüfung und Freischaltung
- Auslieferung der App-Dateien
- Betrieb freigeschalteter optionaler Portale
- Support, soweit die Schule Daten übermittelt

Keine Nutzung für Werbung Dritter. Keine Weitergabe außer an Unterauftragsverarbeiter nach Abschnitt 6.

---

## 5. Technische und organisatorische Maßnahmen (TOM)

Kurzfassung (Details: `docs/sicherheitspruefung-bericht.md`, Trust-Seite):

- Auslieferung über HTTPS; App-CSP mit `script-src 'self'`
- Abhängigkeiten (MSAL, xlsx, jszip, Icons) aus dem eigenen Bundle, nicht von fremden Skript-CDNs in der App
- Deploy über FTPS/SFTP; CI Actions auf Commit-SHA gepinnt
- Login-Scopes je Persona (Portal ohne Admin-Bundle am Anmelden-Button)
- Fachdaten standardmäßig im Schul-Tenant (SharePoint/Graph)

---

## 6. Unterauftragsverarbeiter

| Dienst | Zweck | Region (typisch) |
|---|---|---|
| Microsoft Azure (App Service / Functions / Storage der Betreiber-APIs) | Lizenz- und optionale Portal-Backends | EU (konfiguriert) |
| Hosting-Anbieter für statische Dateien / FTP-Ziel | Auslieferung von app.ms365.schule / ms365.schule | laut Vertrag |

Änderungen teilt der Auftragsverarbeiter der Schule mit; Widerspruch innerhalb von 14 Tagen möglich, soweit gesetzlich vorgesehen.

---

## 7. Betroffenenrechte und Unterstützung

Die Schule bleibt Ansprechpartnerin für Auskunft, Berichtigung, Löschung zu Schulkonten und SharePoint-Daten.  
Der Auftragsverarbeiter unterstützt bei Anfragen, die Betreiber-Minimaldaten betreffen (Lizenzdatensatz, Kontakte).

---

## 8. Löschung und Rückgabe

Nach Ende der Nutzung löscht der Auftragsverarbeiter Lizenz- und Kontaktdaten in den Betreiber-Systemen innerhalb von **90 Tagen**, sofern keine gesetzlichen Aufbewahrungspflichten entgegenstehen.  
Ein Export der Betreiber-Minimaldaten erfolgt auf Anfrage der Schule in maschinenlesbarem Format (z. B. JSON/CSV).

---

## 9. Meldepflichten

Der Auftragsverarbeiter meldet der Schule unverzüglich, sobald ihm eine Verletzung des Schutzes personenbezogener Daten in den von ihm betriebenen Diensten bekannt wird.

---

## 10. Unterschriften

| | Schule | Auftragsverarbeiter |
|---|---|---|
| Ort, Datum | | |
| Name, Funktion | | |
| Unterschrift | | |

---

*Vorlage · MS365-Schul-Tools · 21.11.2026*
