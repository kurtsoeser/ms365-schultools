# Monetarisierung – MS365-Schul-Tools

Stand: 2026-09-22  
Status: **Vorüberlegungen** (noch keine finale Preis-/Produktentscheidung)

Dieses Dokument hält die bisherige Analyse fest: Was das Produkt kann, wie vergleichbare Angebote am Markt positioniert sind, welche Preismodelle sinnvoll sind und welche offenen Fragen vor einem konkreten Angebot noch zu klären sind.

---

## 1. Produktkern (kurz)

Die App ist eine **Browser-Werkzeugkiste für Schul-IT in Microsoft 365** (Fokus AT/DE): Stammdaten lokal, Aktionen über Microsoft Graph / PowerShell.

In fünf Stichworten (Angebots-tauglich):

1. **Gesamte M365-Schulstruktur** – Klassen, Kursteams, Fächer/ARGEs, Sammelgruppen, Diplomarbeiten, Spielwiesen – anlegen, matchen, syncen  
2. **Personen & Zugänge** – Import (CSV, Sokrates, WebUntis, M365), Lizenzen, Gäste, Namenskonventionen, Konten  
3. **Schuljahreswechsel** – geführter Assistent (Backup → Umbenennen → Matura-Jahrgang → Listen → Unterrichtsteams)  
4. **Kommunikation & Intranet** – Shared Mailboxes, Verteiler, Eltern, Bookings, SharePoint-Listen, Power-Automate-Rezepte  
5. **Hygiene & Governance** – Abgleichen, leere Gruppen, Archiv, Policies, Cleanup-Playbook  

Verkaufssatz (Arbeitstitel):  
*Nicht nur Konten anlegen – die komplette M365-Schulwelt: Kursteams, Schuljahr, Hygiene, Kommunikation.*

---

## 2. Marktvergleich (Anker)

| Anbieter | Region | Fokus | Preis (ca., Stand Recherche) | Relevanz |
|----------|--------|--------|------------------------------|----------|
| **LAN.FX** (te.comp) | AT | On-Prem AD-Konten aus CSV (OUs, Home, Quotas); O365 nur am Rande | Schullizenz nach Größe; Einzelpreis oft unklar / im FX-Paket | Schwache Überlappung – eher „User provisioning“ |
| **Virtualschool** | AT | AD + M365-User/Lizenzen + Teams (WebUntis) + Intune + Klassenraum | ~**185 €/Monat ≈ 2.220 €/Jahr** + Setup ~700 € | Starke Überlappung bei User/Teams; **plus Geräte/Server**, die wir (noch) nicht haben |
| **Vis365** (DrVis) | DE | M365-Konten, Lizenzen, Gruppen/Teams, Verteiler, Shared Mailboxes, Eltern-Mails; Premium + Geräte | Standard **600 €** (allg.) / **900 €** (berufsbildend); Premium **2.100 €**/Jahr | Engster DE-Vergleich für reines M365-Admin |
| **Teamsoft AdminTool 2.0** | DE | Konten, Teams, Gruppen, CSV, Schuljahreswechsel | **399 €/Jahr** inkl. MwSt. | Untere Preisgrenze „schul-taugliches M365-Tool“ |
| **School Data Sync (SDS)** | MS | Roster → User/Klassen/Teams | oft **0 €** | Kostenlos, aber kein Schuljahres-/Hygiene-/Automations-Workflow |

### Einordnung unseres Tools

| Wettbewerber | Wir relativ dazu |
|--------------|------------------|
| LAN.FX | Deutlich **mehr M365-Tiefe** (Teams, Struktur, Automationen) |
| Teamsoft / Vis365 Standard | Mindestens **gleichwertig**, in Schuljahr/Kursteams/Hygiene oft **stärker** |
| Virtualschool / Vis365 Premium | Ähnlich bei M365-Betrieb; **schwächer**, wo Intune/Klassenraum/Server mitverkauft werden |

Quellen (Orientierung):  
- https://web.tecomp.at/lanfx.aspx  
- https://virtualschool.at/ / https://www.virtualschool.at/unsere-preise/  
- https://drvis.de/products/vis365 / https://drvis.de/Vis365Pricing  
- https://www.teamsoft.de/microsoft/dienstleistungen/admintool-20-fuer-m365/

---

## 3. Schulbudget-Realität

Schulen zahlen selten nach Feature-Tiefe, sondern nach **Entscheidungsaufwand**:

| Bandbreite / Jahr | Typische Logik |
|-------------------|----------------|
| **unter ~500 €** | oft aus laufendem IT-/Sachmittelbudget, ohne großen Beschluss |
| **~500–1.200 €** | machbar mit kurzer Begründung („spart Schuljahres-Chaos“) |
| **1.500–2.500 €** | eher mit Story, Support, Rahmenvertrag oder wenn ein teureres Tool ersetzt wird |

**Bauchgefühl des Anbieters:** ~**300 €/Jahr** als psychologisch „leicht kaufbar“.  
**Marktanker Vollpakete:** ~**1.800–2.200 €** (Virtualschool, Vis365 Premium) – nicht automatisch, was *jede* Schule spontan zahlt.

Fazit: Hohe Preise sind am Markt belegt, **Einstieg um 300 € ist trotzdem strategisch sinnvoll**, wenn Support und Mehrwert klar gestaffelt sind.

---

## 4. Sinnvolle Monetarisierungsmodelle

### 4.1 Schulpauschale (empfohlen als Grundform)

- Preis **pro Schule / Jahr**, nicht pro User (800 User würden sonst „teuer wirken“)  
- Einfach zu kommunizieren und zu fakturieren  
- Passt zu Admin-Tools (wenige Nutzer der App, Nutzen für die ganze Schule)

### 4.2 Modular / Feature-Verkauf (sehr gut geeignet)

Schulen brauchen oft nur 2–3 Schmerzpunkte. Module verkaufen Features, ohne All-inclusive-Schock.

**Vorschlag Modulzuschnitt (max. 4–6 Module):**

| Modul | Inhalt (Stichworte) | Nachfrage |
|-------|---------------------|-----------|
| **Basis** (immer) | Stammdaten, Dashboard, Import, Personen suchen, Basis-Match | Pflicht |
| **Unterricht & Teams** | Klassengruppen, Kursteams, Vorlagen, ARGEs | hoch |
| **Schuljahr** | Assistent, Umbenennen, Matura-Jahrgang, Archiv | sehr hoch (1×/Jahr) |
| **Personen & Gäste** | Lizenzen, Gäste, Namenskonvention, Sammelgruppen | mittel–hoch |
| **Kommunikation** | Shared Mailboxes, Verteiler, Eltern, Bookings | mittel |
| **Intranet & Automationen** | SharePoint-Listen, Flows, Freistellung, Termine | eher optional |
| **Hygiene & Governance** | Cleanup, leere Gruppen, Policies, Abgleich | für „ordentliche“ Schulen |

Zugpferde: **Schuljahr** + **Kursteams/Unterricht**.

Technisch später: **Feature-Flags / Lizenzschlüssel pro Schule**, nicht getrennte Apps.

### 4.3 Bundle-Staffeln

| Paket | Inhalt | Preisrichtung (Arbeitshypothese) |
|-------|--------|----------------------------------|
| **Starter** | Basis (+ 1 Modul) | **290–390 €**/Jahr |
| **Schulalltag** | Basis + Unterricht + Schuljahr | **390–490 €**/Jahr |
| **Schule** | + Support (E-Mail) | **790–990 €**/Jahr |
| **Alles / Plus** | alle Module + Onboarding | **1.490–1.890 €**/Jahr |

Verkaufssatz: *„Ihr zahlt nur, was ihr nutzt.“*

### 4.4 Setup + niedriges Abo

- Einmalig Setup/Onboarding: **500–900 €**  
- Danach Abo: **300–400 €**/Jahr  

Vorteil: erster Cashflow + niedriger „wiederkehrender“ Preis für Schulen.

### 4.5 Einführungsjahr / Ramp

- Jahr 1: **290 €**  
- ab Jahr 2: **490–790 €**  

Senkt Einstiegshürde, signalisiert aber Wertsteigerung.

### 4.6 Modelle, die vorerst weniger passen

| Modell | Warum vorsichtig |
|--------|------------------|
| **Rein per User** | Admin-Tool; Schulen rechnen dann „800 × X“ und erschrecken |
| **Nur einmal Kauf ohne Updates** | Feature-Fläche wächst; Support/Graph-Änderungen brauchen laufende Pflege |
| **Zu feine Feature-Mikropreise** | Beratungsaufwand, Lizenzchaos |
| **Sofort 2.000 € All-inclusive** | Markt zeigt das bei Vollpaketen – aber viele Schulen springen nicht an |

---

## 5. Preisband – Arbeitsempfehlung (noch nicht final)

Für eine Schule mit ca. **800 User**:

| Strategie | Richtwert / Jahr | Kommentar |
|-----------|------------------|-----------|
| Psychologischer Einstieg | **~300 €** | Bauchgefühl; Teamsoft-Nähe; wenig Support |
| Modularer Sweet Spot | **390–990 €** | Bundle + 1–2 Module |
| Marktvergleich „Volles M365-Ops + Support“ | **1.500–2.200 €** | Vis365 Premium / Virtualschool-Nähe – nur mit klarer Support-Story |

**Aktuelle Haltung:** Modularisierung und Einstieg um **~300 €** im Auge behalten; höhere Preise als Option für Plus/Support, nicht als einziges Angebot.

---

## 6. Was im Angebot kommunizieren (Kernnutzen)

Nicht Feature-Listen-Overload, sondern Outcomes:

1. Schuljahreswechsel **ohne Chaos**  
2. Kursteams/Klassen **aus Stundenplan-Daten**, nicht manuell  
3. Tenant bleibt **ordentlich** (Hygiene, Archiv, Governance)  
4. Weniger Abhängigkeit von Einzel-Skripten / Einzelwissen  
5. Optional: Automationen & Intranet ohne Power-Platform-Projekt von Null  

Abgrenzung zu LAN.FX explizit: *AD-Konten ≠ komplette M365-Schulstruktur.*

---

## 7. Offene Fragen / nächste Überlegungen

### Produkt & Lizenzierung

- [ ] Welche Module sind **Jahr-1-MVP** zum Verkauf?  
- [ ] Feature-Flag-Mechanismus (PIN, Lizenzdatei, Tenant-ID-Whitelist, …)?  
- [ ] Demo-/Musterschule vs. Vollversion – wie lange Trial?  
- [ ] Hosting: weiter GitHub Pages / self-host vs. gehostetes Angebot mit Support-SLA?

### Preis & Vertrag

- [ ] Netto/Brutto, USt, Rechnung als Einzelperson/Firma?  
- [ ] 1-Jahres- vs. 3-Jahres-Rabatt (wie te.comp)?  
- [ ] Was ist im Starter-Preis an Support enthalten (Antwortzeit, Stunden/Jahr)?  
- [ ] Bildungsdirektion / Rahmenvertrag später interessant (Virtualschool-Vorbild Tirol/Salzburg)?

### Wettbewerb & Positionierung

- [ ] Primär AT, DE, oder beides? (Import: Sokrates/WebUntis vs. ASV/Schild)  
- [ ] Bewusst **ohne Intune** bleiben oder später Geräte-Modul?  
- [ ] Konkurrenz-Story: „günstiger Einstieg als Virtualschool, tiefer als LAN.FX“

### Rechtliches / Betrieb

- [ ] AGB, AVV (falls je Server-seitige Daten), Datenschutztext an kommerzielle Nutzung anpassen  
- [ ] Support-Kanal und Erreichbarkeit festlegen, bevor Plus verkauft wird  

---

## 8. Kurzfazit

1. Das Tool ist **marktüblich monitarisierbar** – vergleichbare Produkte liegen zwischen **~400 €** und **~2.200 €**/Jahr.  
2. Viele Schulen wollen eher **~300 € Einstieg** als Vollpreis – das ist kein Widerspruch, sondern ein **Staffel-/Modul-Thema**.  
3. **Schulpauschale + Module + Bundle** ist das sinnvollste Modell.  
4. Höhere Preise erst mit **Support/Onboarding** und klarer Abgrenzung zu LAN.FX / SDS verkaufen.  
5. Nächster Schritt, wenn es konkret wird: MVP-Module festnageln + eine Preistabelle „Starter / Schulalltag / Plus“ finalisieren.

---

## Anhang A – grobe Feature→Modul-Zuordnung

Zur späteren technischen Feature-Flag-Planung (nicht verbindlich):

| Bereich in der App | Modul-Kandidat |
|--------------------|----------------|
| Stammdaten, tenant, Import, Dashboard | Basis |
| Klassengruppen, Kursteams, Vorlagen, ARGEs, Diplom, Spielwiesen | Unterricht & Teams |
| organisations-Assistent / Schuljahr, Archiv-Jahrgang | Schuljahr |
| Personen-Verwaltung, Gäste, Namenskonvention, Sammelgruppen | Personen & Gäste |
| Shared Mailboxes, Verteiler, Eltern, Bookings | Kommunikation |
| Intranet, Listen, Power-Automate-Rezepte | Intranet & Automationen |
| Cleanup, leere Gruppen, Policies, Datenhygiene, Abgleich | Hygiene & Governance |

## Anhang B – Änderungslog dieses Dokuments

| Datum | Änderung |
|-------|----------|
| 2026-09-22 | Erstfassung aus Analyse + Marktvergleich (LAN.FX, Virtualschool, Vis365, Teamsoft) und interner Preisdiskussion (~300 € vs. Marktanker) |
