# Stammdaten Lokal-First (Umbau)

**Status:** aktiv (`STAMMDATEN_LOCAL_FIRST = true` in [`src/shared/stammdaten-local-first.js`](../src/shared/stammdaten-local-first.js))  
**Phase 4:** manuelle IT-Bibliothek freigeschaltet (`STAMMDATEN_IT_MANUAL_BACKUP = true`)  
**Phase 5:** manueller Intranet-Listen-Export freigeschaltet (`STAMMDATEN_INTRANET_MANUAL_EXPORT = true`)  
**Verwandt:** [schul-struktur-kanon.md](schul-struktur-kanon.md), [stammdaten-v1-sunset.md](stammdaten-v1-sunset.md), **[daten-inventar-landkarte.md](daten-inventar-landkarte.md)** (wo welche Daten liegen)

## Live-Listen vs. JSON-Backup

| Schicht | Zweck |
|---------|--------|
| **SharePoint-Listen** (Lehrer, Klassen, Schulaktivitäten, …) | Betriebs-Wahrheit im Schul-Tenant; in SP sichtbar/pflegbar; App liest/schreibt hier |
| **IT-JSON-Backup** (`ms365-stammdaten-aktuell.json`) | Absicherung / Bootstrap leerer Browser – **kein** Ersatz für Live-Pflege |
| **Browser localStorage** | Arbeitscache, nie schulweite Wahrheit |

Beim **Schulaktivitäten-Listen-Paket** werden Lehrer- und Klassen-Listen mit Schema mitgeprüft/angelegt (`school-stammdaten-spo-lists.js`). Zeilen kommen aus SP-Pflege oder Stammdaten-Sync – nicht nur aus einem einmaligen JSON.

**Klassen ↔ M365-Gruppe:** Die SharePoint-Klassenliste führt Textfelder `M365GroupId` / `M365GroupMail` / `M365GroupName` sowie das Personen-/Gruppenfeld `M365Gruppe` (peopleAndGroups). Beim Befüllen werden verknüpfte Klassen-Teams (`core.classTeams` / `classGroupMatchByKey`) mitgeschrieben. Der Schulaktivitäten-Planer lädt die Gruppen-Mail für Outlook-Einladungen; fehlt sie, werden vorhandene Schüler-E-Mails der Klasse einzeln eingeladen.

**Apps lesen Live:** Schulaktivitäten-Planer (`refreshData`) und sRDP/Matura (`loadFromSharePoint`) laden Lehrer/Klassen über `fetchSchoolStammdatenFromSite` / `hydrateStammdatenFromLiveSite`. Eine Änderung in der SharePoint-Klassenliste wirkt nach „Laden“ ohne neues JSON-Backup. LocalStorage bleibt nur Fallback, wenn die Listen fehlen oder leer sind.

**Rechte beim Listen-Paket:** Unter `tools/sharepoint-liste-schulaktivitaeten.html` können Verwaltung (Gruppe / mail-aktivierte Sicherheitsgruppe / Einzelperson) und Lehrkräfte gewählt werden. Beim „Paket anlegen“ bzw. „Nur Berechtigungen“ setzt die App Listenrechte (Vererbung brechen): Lehrer → Anträge *Beitragen*, Klassen/Lehrer/Regelwerk *Lesen*; Verwaltung → *Vollzugriff*. Voraussetzung: ausführender Account mit SharePoint-Vollzugriff auf der Site (`Sites.FullControl.All`).

**Befüllung beim Paket:** „Paket anlegen“ bzw. „Nur Stammdaten befüllen“ schreibt die lokalen Stammdaten (Lehrerinnen + Klassen inkl. KV) in die SharePoint-Listen (`seedSchoolStammdatenListsFromLocal`). Schema-only reicht nicht – ohne diesen Schritt bleiben die Listen leer und der Planer hat keine Klassenauswahl.

**Häufiger Fehlgriff:** „Nur Berechtigungen“ setzt ACL, schreibt aber **keine** Zeilen. Health meldet leere Klassen-/Lehrerlisten jetzt als Fehler (`0 Einträge`). Im Planer erscheinen Hinweise, wenn Live-Klassen fehlen.

**Lehrer-Browser ohne IT-Kontext (R13):** Domain/Site/Gruppen liegen local-first beim IT-Admin. Ohne **ToolSiteUrl** in der Lizenz und ohne IT-Sicherung auf SharePoint zeigt ein frischer Lehrer-Login oft leere Domain, „Noch nicht zugeordnet“ und keine Klassen.

- **Schule nicht freigeschaltet:** License-Gate mit Hinweis an IT-Admin + CTA **Anfrage stellen** → [ms365.schule/#kontakt](https://ms365.schule/#kontakt).
- **Freigeschaltet, Einrichtung unvollständig:** dezenter Banner im Lehrer-Portal („Portal bereit?“ in `lehrer-portal-readiness.js`) – IT soll Tool-Site, Listen und IT-Sicherung nachziehen; Lehrer-Tour startet erst, wenn das Portal bereit ist (Menü-Start bleibt möglich).
- **Fix für IT:** Admin → Schule → Tool-Site-URL setzen, IT-Sicherung ablegen, Listen befüllen; Lehrkraft neu anmelden (Bootstrap übernimmt die URL).

## Schul-Site-Zeiger (Lizenz)

Der Betreiber speichert **keine** Klassen-/Lehrer-Inhalte, sondern nur einen URL-Zeiger je Schule:

- Lizenzliste-Feld **`ToolSiteUrl`** (Admin → Schildetail → Lizenz)
- App (`/license/me`) liefert `sites.toolSiteUrl` / `license.toolSiteUrl`
- Nach Login: Browser übernimmt die URL und lädt bei **leerem** Stammdaten-Stand die IT-Sicherung (`ms365-stammdaten-aktuell.json`) **aus der Schul-SharePoint** mit dem User-Token der Schule

**Spalte fehlt in SharePoint:** Wenn die Tool-Site-URL nach „Speichern“ wieder leer ist, existiert `ToolSiteUrl` noch nicht in „Übersicht-Tenants“. Einmal Schema aktualisieren (`ensure-schema` / Script / Admin-Setup), dann URL erneut speichern – intern heißt die Spalte **`ToolSiteUrl`**.

Damit funktioniert der erste Besuch auf einem neuen PC ohne lokale Demo-Daten. Die Wahrheit bleibt im Schul-Tenant.

## Ziel

Eine Schreibwahrheit lokal im Browser, ohne Sync-Chaos:

| Schicht | Inhalt |
|---------|--------|
| **Listen** (`core.*`, `years.*`) | Lehrer, ARGEs, Klassen, Schüler, Belegung, … |
| **Landkarte** (`core.schulOrganigramm`) | Hierarchie, Position, Referenzen, M365-Bindung |

Jeder Speichern-/Import-Pfad schreibt Listen über [`stammdaten-write.js`](../src/shared/stammdaten-write.js) (`writeSchoolData`) und materialisiert die Landkarte – **kein** manueller Abgleich Listen ↔ Landkarte im Alltag.

**Semantik:** Listen = Entitäten-Wahrheit; Landkarte = Struktur-/Layout-/M365-Wahrheit. Organigramm-Editoren schreiben Listen nur punktuell zurück (kein voller Reverse-Materializer). Details: [schul-organigramm.md §8.1](schul-organigramm.md), [schul-struktur-kanon.md](schul-struktur-kanon.md).

## Was aus bleibt

- Auto-Push/Pull SharePoint IT-Bibliothek
- Konflikt-/Abgleich-Dialoge („Lokal vs. IT prüfen“)

## Was erlaubt ist

- **Speichern** lokal (`ms365-schooltool-data-v2`)
- **Browser-Backup** Export/Import (JSON)
- **IT-Sicherungsbibliothek manuell**
  - Einrichten (Stammdaten / Übergabe-Seite)
  - Sichern = aktuelles Browser-Backup nach SharePoint (+ datierte Version)
  - Laden = Backup-Datei aus der Bibliothek wählen und importieren
- **Intranet-Listen manuell** (z. B. „Intranet: Lehrerinnen“ unter `/tenant.html#lehrer`)
  - Schreibt die lokalen Stammdaten in die SharePoint-Liste auf der Intranet-Site
  - Damit Schultools (Schulaktivitäten, sRDP, …) die aktualisierte Live-Liste lesen
- Landkarte und Listen-Tabs pflegen
- Importe (WebUntis / CSV) → schreiben Listen + materialisieren Landkarte

## Flags

| Flag | Bedeutung |
|------|-----------|
| `STAMMDATEN_LOCAL_FIRST` | Auto-Sync gesperrt; kanonisch lokal im Browser |
| `STAMMDATEN_IT_MANUAL_BACKUP` | Manuelles Einrichten/Sichern/Laden der IT-Bib erlaubt |
| `STAMMDATEN_INTRANET_MANUAL_EXPORT` | Manueller Intranet-Listen-Export (Lehrer, Klassen, …) erlaubt |

Vollständigen Auto-Sync erst freigeben, wenn `STAMMDATEN_LOCAL_FIRST = false`.

## Entwickler-Hinweise

- Nicht auf `ms365-tenant-settings-v1` schreiben (Sunset).
- Tool-Nebenstates: [`stammdaten-side-state-keys.js`](../src/shared/stammdaten-side-state-keys.js) – Cache, nicht Wahrheit.
- `ms365TenantSettingsSave` → v2; Materializer hängt an `setCoreFromTenantSettings`.
