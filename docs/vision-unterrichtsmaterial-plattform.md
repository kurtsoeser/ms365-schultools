# Vision: Unterrichtsmaterial-Plattform (OneNote + Kursteams)

**Stand:** 2026-09-23  
**Status:** Ideensammlung / Architektur-Klarstellung – nach dem Durchbruch zu Cross-Tenant & 1:1-Kopie  
**Kontext:** Aufbau auf dem bestehenden MS365-Schultools (Admin-Werkzeugkiste) + zentralem Katalog auf kurtrocks (`MS365-Katalog`: Vorlagen, `materialien`, `notebooks`)

Verwandte Docs: [Monetarisierung](./monetarisierung.md) · [Projektanalyse](./projektanalyse-und-werkzeugideen.md)

---

## 1. Der Durchbruch (Kernidee)

Nicht nur **Schul-IT/Admins** bedienen die Schultools – sondern **einzelne Lehrkräfte** melden sich an und nutzen genau die Dinge, die im Unterricht zählen:

| Für Lehrkräfte relevant | Für Lehrkräfte eher irrelevant |
|-------------------------|--------------------------------|
| OneNote-Inhalte in Kursnotizbücher holen | Mandanten-Policies, SharePoint-Tenant-Settings |
| Kursteam-Kanal-Vorlagen anwenden | Cleanup-Playbook, leere Gruppen |
| Dateien/Materialien aus dem zentralen Fundus | Schuljahreswechsel-Assistent, Namenskonvention-Audit |
| Best-Practice-Strukturen übernehmen | Gäste-Governance, Verteilerlisten-Setup |

**Zielbild:** Eine **Plattform / Tauschbörse für Unterrichtsmaterial und Best Practices** auf Basis von Microsoft 365 – vor allem **OneNote Kursnotizbücher** (komplexe Seiten inkl. Medien, Ink, Tabellen) – aus der sich Lehrkräfte **bequem und gezielt** Inhalte **1:1** in **ihre** Kursteams / Kursnotizbücher ziehen.

Optional später: **gegen Bezahlung** (Schule, Fachgruppe, Einzellizenz, Marketplace).

Das ist der Hebel, der seit Jahren fehlt: nicht „noch ein Download-Portal mit PDFs“, sondern **direkt ins laufende Kursnotizbuch** – als echte OneNote-Kopie, nicht als HTML-Behelf.

---

## 2. Was heute schon die Basis legt

| Baustein | Stand (Repo / kurtrocks) | Rolle in der Vision |
|----------|--------------------------|---------------------|
| Zentraler SharePoint-Katalog | Site `MS365-Schultools`, Bibliothek `MS365-Katalog` | „Fundus“ + Metadaten |
| Kanal-Vorlagen | `vorlagen/…`, Admin pflegt, Schulen lesen | Best-Practice Team-Struktur |
| Materialien | Ordner `materialien/` | Dateien in Kanäle |
| OneNote-Vorlagen | Echte Notizbücher auf der Site (z. B. `MS365-Vorlagen-Notizbuch`) | Quelle für 1:1-Copy |
| Werkzeug OneNote verteilen | `tools/onenote-verteilung.html` | Lehrkraft-tauglicher Einstieg (Prototyp) |
| License-API | Schulen lizenziert, Katalog-Metadaten / Dateien ohne Direktzugriff | Multi-Tenant-Gate |
| Snapshot (HTML/JSON) | `onenote-snapshot/` | Fallback: Text/Tabellen + **eingebettete Bilder** (wo Graph sie liefert); Forms/Learning Activities = Platzhalter |

---

## 3. Architektur-Klarstellung: 1:1-Kopie vs. Snapshot (2026-09-23)

### 3.1 Was Microsoft vorgibt

- **OneNote Graph mit App-Only** ist seit März 2025 für diese APIs **abgeschaltet**.  
  → Eine zentrale Backend-App kann Vorlagen-Notizbücher **nicht** live für alle Schulen auslesen und „durchreichen“.
- **1:1-Kopie** (Abschnitte inkl. Medien, Ink, komplexe Seiten) geht zuverlässig über  
  **`copyToSectionGroup` / vergleichbare Copy-Operationen mit dem Token der Lehrkraft** – also **delegiert**, wenn sie **Quelle und Ziel** sehen darf.

### 3.2 Was damit „tot“ ist – und was nicht

| Idee | Urteil |
|------|--------|
| Komplett **anonym / öffentlicher Link** + Graph kopiert 1:1 inkl. Medien | praktisch **nein** (kein brauchbares User-Token für Copy) |
| Freigabe-Link „Jeder mit Link“ → Schul-User in anderem Tenant öffnet als **Gastmitwirkender** | UX/Identität bricht; Graph mit Schul-Login sieht oft **keinen** Zugriff |
| Gezielte Gast-Einladung B2B kurtrocks ↔ Schul-Tenant | **theoretisch** möglich, in der Praxis für OneNote-Site-Notizbücher **sehr schwer** (Einladung, Konto-Konflikt, OneNote-Web, Rechte) |
| HTML-Snapshot (Seiten als HTML neu anlegen) | **Notnagel** – Text/Tabellen + Bilder (beim Veröffentlichen eingebettet); **kein** 1:1 bei Forms, Learning Activities, Stream, Ink |
| **1:1 nur im gleichen Tenant** (kurtrocks / Schul-IT mit Site-Zugriff) | **zuverlässig** |
| Tauschbörse mit Login + Metadaten + Materialien/Kanäle; OneNote-1:1 wo Tenant-Zugriff da ist | **ja – produktfähig** |

**Erfahrung 2026-09-23:** Notizbücher, die auf einem MS365-Schul-Tenant (kurtrocks) liegen, lassen sich für User in **anderen** Schul-Tenants kaum so freigeben, dass automatisiert und zuverlässig **1:1** (Medien, Ink) per Graph kopiert werden kann. Gastmitwirkender-Links ≠ Schul-Konto in der App. Das ist ein **Microsoft-Plattform-Limit**, kein reines UI-Problem der Schultools.

Die Tauschbörse-Idee ist deshalb **nicht tot**, aber das OneNote-Versprechen muss ehrlich sein:

> **Cross-Tenant 1:1-OneNote ist kein Standard-Feature**, das wir „mit einem Freigabe-Link“ liefern können.  
> Kern der Börse: Katalog, Lizenz, Kanäle, Dateien, Vorschau; OneNote-1:1 dort, wo echter Tenant-/Site-Zugriff besteht; sonst Snapshot oder manueller OneNote-Weg.

### 3.3 Zielmodell für die OneNote-Tauschbörse

```
[Katalog / Börse]          Metadaten, Vorschau, Suche, Lizenzcheck
        │                  (License-API – App-Only auf SharePoint-Dateien/Listen ok)
        ▼
[Freigabe der Vorlage]     Lehrkraft darf das Quell-Notizbuch lesen
        │                  (Gast / gezielte Freigabe / Verbund-Zugriff – nicht „anonym öffentlich“)
        ▼
[1:1 Copy]                 Schul-Token: copyToSectionGroup → eigenes Kursnotizbuch
                           (Medien, Ink, komplexe Seiten bleiben OneNote-nativ)
```

**Rollen der Bausteine:**

1. **Katalog (License-API)** – Was gibt es? Fach, Stufe, Autor, Version, Status „für meine Schule freigeschaltet?“  
2. **Zugriff auf die Vorlage** – Einmalige oder schulweite Freigabe der echten OneNote-Notizbücher (kurtrocks oder Autor:innen-Tenant) an die lizenzierten Schul-Konten / Gäste.  
3. **Verteilen** – Tool nutzt den **Schul-User**, nicht die Betreiber-App, für den Copy.

**Snapshot** bleibt optional für Notfälle (kein Freigabe-Weg möglich). Beim Veröffentlichen werden OneNote-Bildressourcen (und kleine Dateianhänge) per Publisher-Token eingebettet; **Forms / Learning Activities / Stream** werden bewusst als sichtbarer Platzhalter markiert – die lassen sich über Graph nicht sinnvoll „mitkopieren“. Nie als Qualitätsversprechen „1:1 inkl. interaktiver Embeds“.

### 3.4 Freigabe-Varianten (produktseitig)

| Variante | Aufwand | 1:1 möglich? | Eignung |
|----------|---------|------------|---------|
| B2B-Gast / Freigabe an Schul-UPNs | mittel | ja | Workshop, Pilotschulen |
| Schulweite Gruppe / Verbund-Zugriff auf Vorlagen-Site | höher, skalierbar | ja | Lizenzkunden |
| „Anyone with the link“ anonym | niedrig | Graph-Copy eher nein | nur manuell in OneNote öffnen |
| HTML-Snapshot | schon gebaut | eingeschränkt | Fallback |

---

## 4. Zwei Produkt-Ebenen (klar trennen)

### A) Betrieb / Admin („Schultools IT“)

Unverändert wichtig: Stammdaten, Klassen, Kursteams anlegen, Schuljahr, Hygiene, Policies.  
Zielgruppe: **IT, Schulleitung, Power-User.**

### B) Unterricht / Lehrkraft („Schultools Unterricht“ / Tauschbörse)

Schlanke Oberfläche nur mit:

1. **OneNote-Vorlagen** → 1:1 in meine Kursnotizbücher (nach Freigabe)  
2. **Kanal-Vorlagen** → auf meine Kursteams  
3. **Material-Bibliothek** → Dateien in Kanäle  
4. Optional: **Meine Favoriten**, **Fachfilter**, **Schulform/Schulstufe**

Dashboard und Navigation: getrennte „Welt“ oder Rollenfilter („Ich bin Lehrkraft“ vs. „Ich bin Admin“).

---

## 5. Community & Marketplace – Ideen

### 5.1 Lesen (Konsum)

- Katalog nach **Fach, Schulform, Schulstufe, Semester, Schlagworten**  
- Vorschau: Abschnittsliste, Seitenanzahl, Beschreibung, Autor, Aktualisierungsdatum  
- Klarer Status: **Freigabe aktiv** / **Freigabe anfordern** / nur Snapshot-Fallback  
- „In meine Kursnotizbücher“ (ein Team oder Verteilerliste) – **1:1 Copy**  
- Bewertungen / „an meiner Schule genutzt“

### 5.2 Einreichen (Community)

- Lehrkräfte / Fachgruppen reichen **Vorlagen-Pakete** ein:
  - OneNote-Abschnitte (aus eigenem Vorlagen-Notizbuch – nach Freigabe für den Börsen-Betrieb)
  - Kanal-Sets (JSON wie heute)
  - Material-Ordner (PDF, Arbeitsblätter)
- Redaktion / Moderator (Betreiber oder Schulverbund) freigibt  
- Versionierung: v1.2, Changelog, „kompatibel mit Schuljahr …“  
- Attribution: Name, Schule, Lizenz (CC-BY, nur innerhalb MS365-Schultools, kommerziell, …)

### 5.3 Kuratierte Best Practices

- „Offizielle“ Pakete vom Betreiber (z. B. HAK BEFC Stufe 10)  
- Schulinterne Kataloge (nur Tenant X sieht Ordner `schule-xyz/`)  
- Verbünde / Bildungsdirektion: gemeinsamer Fundus + gemeinsame Freigabe-Regel

### 5.4 Bezahlung (Anknüpfung an Monetarisierungs-Doc)

Mögliche Modelle (nur Ideen):

| Modell | Beschreibung |
|--------|--------------|
| Schule flat | Alle Lehrkräfte der lizenzierten Schule nutzen den Fundus + Freigabe |
| Fach-Paket | z. B. nur Mathematik / nur HAK |
| Autor:innen-Share | Community-Vorlagen mit Umsatzbeteiligung |
| Freemium | Basis-Vorlagen gratis, Premium-Pakete kostenpflichtig |
| Pro Lehrkraft | Kleinpreis für Quereinsteiger ohne Schulvertrag |

**Content-Freigabe folgt der Lizenz:** Lizenz aktiv → Schule darf Vorlagen sehen und (nach Freigabe-Schritt) 1:1 kopieren.  
Wichtig: **Admin-Produkt** und **Content-Produkt** können getrennte Preise haben.

---

## 6. Weitere Möglichkeiten (Backlog Vision)

### UX / Produkt

- **Lehrkraft-Home:** „Meine Kursteams“ aus Graph – Vorlagen mit einem Klick auf alle Mathe-Teams  
- **Diff & Update:** „Vorlage hat neue Version – Abschnitte nachziehen?“ (ohne Schülerseiten zu zerstören)  
- **Nur Inhaltsbibliothek** als Default; Lehrerbereich getrennt  
- **Vorlagen-Wizard:** „Neues Schuljahr → empfohlene Pakete“  
- **Teams-App / Tab** später  
- Freigabe-Assistent: „Meine Schule freischalten“ (ein Klick → Gast/Freigabe-Flow)

### Inhaltstypen

- OneNote-Abschnitte / Abschnittsgruppen (**1:1 bevorzugt**)  
- Kanalstrukturen  
- Dateien (Arbeitsblatt, Rubric, Checkliste)  
- Später: Loop, Assignments, Forms als Metadaten  
- „Klassennotizbuch-Starter“

### Technik / Vertrauen

- **Katalog/Metadaten/Dateien:** License-API (App-Only auf SharePoint ok)  
- **OneNote 1:1:** immer **Schul-Token** + Leserecht auf Quelle + Schreibrecht auf Ziel  
- **Cross-Tenant:** nicht „App liest kurtrocks für alle“, sondern „Schule bekommt Zugriff, User kopiert selbst“  
- HTML-Snapshot: dokumentierter Fallback – Bilder beim Publish einbetten; Forms/Learning Activities bewusst Platzhalter  
- Audit: wer hat wann welche Vorlage in welches Team gespielt  
- DSGVO: keine Schülerdaten im zentralen Katalog

### Community-Prozesse

- Einreich-Formular + Checkliste (Barrierefreiheit, Bildrechte, …)  
- Peer-Review, Flags bei veralteten Lehrplanbezügen  
- Schilf / Hackathons: gemeinsam Vorlagen bauen

### Partnerschaften

- Schulbuchverlage, ARGE, Microsoft Education (Showcase Class Notebook + zentrale Vorlagen)

---

## 7. Risiken & offene Fragen

1. **Qualitätssicherung** – wer haftet für inhaltliche Fehler?  
2. **Lehrplan-Bezug** – AT Bund/Länder; Metadaten müssen das abbilden  
3. **Rechte an Inhalten** – klare Nutzungsbedingungen (keine Schulbuch-Scans ohne Recht)  
4. **Freigabe-Skalierung** – wie laden wir hunderte Schul-UPNs / Gäste ohne Chaos? (Gruppen, Automatisierung, Partner-Center)  
5. **Erwartung „öffentlicher Link = 1:1 API“** – klar kommunizieren: Login + Freigabe nötig  
6. **Admin vs. Lehrkraft** – möglichst wenig Graph-Scopes (`Notes.*`, Teams lesen)  
7. **Fallback Snapshot** – wann akzeptabel, wann irreführend?

---

## 8. Grobe Phasen (Orientierung, aktualisiert)

| Phase | Ziel |
|-------|------|
| **Jetzt / Prototyp** | kurtrocks: echte Notizbücher + Verteilen (Site-Graph); Snapshot nur Fallback |
| **P1 Lehrkraft-UI** | Reduzierter Einstieg, „meine Teams“, klare Trennung Admin/Unterricht |
| **P2 Freigabe + 1:1 Multi-Tenant** | Lizenzierte Schule bekommt Vorlagen-Zugriff; Copy per Schul-Token |
| **P3 Tauschbörse / Community** | Einreichen, Freigabe-Workflow, Versionen, Attribution |
| **P4 Marketplace** | Bezahlmodelle, Freemium, optional Autor:innen-Share |

---

## 9. Merksatz (für später)

> **MS365-Schultools wird von der Admin-Werkzeugkiste zur Unterrichts-Tauschbörse:**  
> kuratierte und community-getragene OneNote-/Kanal-/Datei-Vorlagen –  
> Lehrkräfte holen Best Practices **1:1** (inkl. Medien) mit wenigen Klicks in ihre Kursnotizbücher.  
> Katalog und Lizenz steuern **wer darf**; Freigabe + User-Token steuern **wie OneNote kopiert**.  
> App-Only-Snapshots sind nur der Notnagel – nicht das Produktversprechen.

---

## 10. Schnelle Notizen

### 2026-09-22 (Roh)

- Einloggen einzelner Schul-User (nicht nur Admins)  
- Nur Tools: OneNote, Kursteam-Vorlagen, Dateien/Materialien  
- Plattform für Unterrichtsmaterial + Best Practice  
- Bequem aus dem Fundus in die eigenen Kursnotizbücher  
- Evtl. Bezahlung · Community einreichen/teilen  

### 2026-09-23 (Architektur)

- OneNote Graph App-Only tot → kein „Backend liest kurtrocks für alle Schulen live“  
- 1:1 inkl. Medien/Ink = delegierter Copy – **nur zuverlässig mit echtem Zugriff im Quell-Tenant**  
- Öffentlicher Link / Gastmitwirkender ≠ Schul-UPN in der App → Cross-Tenant-Freigabe von Site-Notizbüchern **praktisch sehr schwer**  
- Tauschbörse lebt mit ehrlichem Scope: Katalog + Lizenz + Kanäle/Dateien; OneNote-1:1 nicht als Cross-Tenant-Standard verkaufen  
- HTML-Snapshot = Fallback, nicht Kernversprechen  

*Dieses Dokument darf und soll erweitert werden, sobald erste Nutzer:innen-Feedbacks oder Preisentscheidungen vorliegen.*
