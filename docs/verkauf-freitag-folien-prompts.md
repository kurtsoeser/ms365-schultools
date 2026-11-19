# Copilot-PowerPoint-Prompts – Freitag Workshop Tag 3

**Zweck:** Folien als **Begleitung** zum Hands-on (nicht als Vortrag).  
**Vorbild-Stil:** `2026-11-06 - MS365 für die Schulverwaltung - BLOCK B.pptx`  
**Inhaltlich:** `docs/verkauf-freitag-workshop.md` (Tag-3-Playbook)  
**Datum/Zeit (anpassen falls nötig):** Freitag, 21. November 2026 · 09:30–16:15  
**Marke:** MS365-Schultools · #kurtrocks edu.innovation.consulting · ms365.schule

---

## So nutzt du die Prompts in Copilot PowerPoint

1. Neue Präsentation öffnen → Copilot → zuerst den **Master-Prompt (Stil)** ausführen.  
2. Danach entweder:
   - **Variante A:** den **Gesamtdeck-Prompt** einmal senden, oder  
   - **Variante B (empfohlen):** Folien **blockweise** mit den Abschnitts-Prompts erzeugen (bessere Kontrolle).  
3. Am Ende einzelne Folien mit den **Feinschliff-Prompts** nachziehen.  
4. Keine Screenshots der App auf die Folien – Live-Demo läuft im Browser; Folien steuern nur den Rahmen.

---

## 0) Master-Prompt – Stil & Regeln (immer zuerst)

```text
Erstelle eine professionelle Workshop-Präsentation auf Deutsch im Stil einer Moderatoren-Begleitung (kein Marketing-Pitch, keine lange Firmenvorstellung).

Stil-Vorgaben (wie bei einem klaren Schulungs-Workshop-Deck):
- Klare, große Überschriften; wenig Fließtext; viele kurze Bullet-Punkte
- Zeitblöcke mit Uhrzeit, Format (Plenum / Schulteam / Demo) und einem Kasten „ERGEBNIS“
- Nummerierte Schritte 01, 02, 03, 04
- Seitenkästen mit Leitfragen (z. B. „ACHTET AUF“, „ROLLEN“, „VERSPRECHEN“)
- Dezente Pausen-Folien nur mit dem Wort PAUSE und Uhrzeit
- Farbwelt: ruhig, schulisch-professionell, hoher Kontrast, keine verspielten Illustrationen, keine Emoji-Flut
- Marke unten klein: „MS365-Schultools · ms365.schule · #kurtrocks“
- Keine Produktkatalog-Folien mit 20 Tools; Fokus auf Prozess, Rollen, Mitnahmen
- Titel der Veranstaltung: „MS365 für die Schulverwaltung – TAG 3“
- Untertitel: „Bauen · Mitnehmen · Entscheiden“
- Zielgruppe: Direktion, Administration, IT/Admin (bereits bekannt aus Termin 1 und 2)

Inhaltliche Haltung auf allen Folien:
- Das ist der 3. Workshop-Tag – kein Vorstellen, kein Vision-Vortrag
- Teilnehmende haben Use Cases und Feedback mitentwickelt
- Moderator = Übermittler; Prinzip auch mit SharePoint/Forms/Approvals möglich
- Heute mindestens 3 funktionsfähige Dinge mitnehmen
- Am Ende: 2-Wochen-Probe, dann ruhige Entscheidung – kein Kaufdruck auf den Folien
```

---

## 1) Gesamtdeck-Prompt (wenn Copilot alles auf einmal bauen soll)

```text
Erstelle eine komplette Workshop-Präsentation (~28–32 Folien) auf Deutsch für:

Titel: MS365 für die Schulverwaltung – TAG 3
Untertitel: Bauen · Mitnehmen · Entscheiden
Zeit: Freitag, 09:30–16:15
Kontext: Dritter Workshop-Tag einer Reihe. Die Schulen kennen sich und den Moderator. Feedback der letzten Wochen steckt bereits in den Lösungen. Heute wird im eigenen Microsoft-365-Tenant gebaut.

Deck-Struktur exakt in dieser Reihenfolge:

A) Eröffnung (Folien 1–8)
1. Titelfolie: MS365 für die Schulverwaltung | TAG 3
2. Datum/Zeit: 09:30–16:15 | Freitag
3. Branding-Folie: #kurtrocks edu.innovation.consulting
4. Kurzlink-/URL-Folie: ms365.schule (App: app.ms365.schule) – groß, zentriert
5. „Unser Tag von 09:30 bis 16:15“ – Agenda als Zeitstrahl mit 7 Blöcken
6. Haltung: „Kein Vortrag – wir bauen.“ Drei Sätze: mitentwickelt / Übermittler / 3 Mitnahmen
7. Rollen-Wow: 3 Spalten Direktion | Administration | IT – je 1 Nutzen
8. Versprechen: „Heute nehmt ihr mindestens 3 lauffähige Dinge mit“ – Liste der 3 Mitnahmen

B) Block 1 – Rein in die App (Folien 9–12)
9. Abschnittstitel: 1 | 09:40–10:10 | Login & Consent | ERGEBNIS: Jede Schule ist angemeldet
10. Login-Checkliste 01–06 (URL, Schulkonto, MFA, Admin-Consent, neu anmelden, Hilfe)
11. Consent in Alltagssprache: Daten bleiben im Schul-Tenant; Admin-Zustimmung einmalig nötig
12. Dealbreaker: Global Admin / Application Admin muss erreichbar sein

C) Block 2 – Stammdaten (Folien 13–15)
13. Abschnittstitel: 2 | 10:10–11:20 | Stammdaten gemeinsam | ERGEBNIS: Klassen + Lehrkräfte + Site
14. Minimum heute: Klassen-Codes, Lehrkräfte mit E-Mail, Site/Listen-Ort
15. Co-Building-Satz: „Wir füllen eure Basis – nicht meine Demo-Daten.“ Hinweis WebUntis/Vorlagen

D) Block 3 – Mitnahme 1 Schulaktivitäten (Folien 16–19)
16. Abschnittstitel: 3 | 11:20–12:15 | Mitnahme 1: Schulaktivitäten | ERGEBNIS: 1 kompletter Antrag-Flow
17. Ablauf 01 Melden 02 Prüfen 03 Entscheiden 04 Informieren/Kalender 05 In SharePoint speichern
18. ACHTET AUF: Was überzeugt? Was fehlt? Welcher Sonderfall?
19. Strategie: „Wahrheit liegt in SharePoint – Alltag auch ohne Dauer-App-Nutzung.“ Scope: keine Auto-Untis

PAUSE-Folie: 12:15 Mittag

E) Block 4 – Mitnahme 2 (Folien 20–23)
20. Abschnittstitel: 4 | 13:00–13:45 | Mitnahme 2 | ERGEBNIS: zweiter lauffähiger Mini-Flow
21. Wahlfolie: Lehrerfreistellung („ich möchte frei haben“) ODER Matura/sRDP („nach Weihnachten“)
22. Lehrerfreistellung kurz: Antrag → Direktion → Kalender/iCal | Abgrenzung Schüler-Freistellung
23. Matura/sRDP kurz: Anmeldung → Lehrerblick → Auswertung | HAK priorisieren

F) Block 5 – Mitnahme 3 Demo-Hebel (Folien 24–26)
24. Abschnittstitel: 5 | 13:45–14:05 | Mitnahme 3: Kalender + Kursteams-Demo | ERGEBNIS: IT-Hebel sichtbar
25. Kalender-Nutzen: ICS/Feed aus dem heutigen Prozess – Admin/Direktion
26. Kursteams-Demo: 8–12 Minuten, bewusst Demo – „so einfach entlastet IT“

G) Block 6 – Individuell (Folien 27–28)
27. Abschnittstitel: 6 | 14:05–15:35 | Individuell / Vertiefung | ERGEBNIS: schulischer Use Case vertieft
28. Arbeitsmodus: Moderator wandert; Schulen an eigenem Fall; 3er-Liste abhaken

H) Close (Folien 29–32)
29. Abschnittstitel: 7 | 15:35–16:15 | Feedback & Close | ERGEBNIS: Owner + Termin in ~2 Wochen
30. Feedback-Würdigung: „Euer Feedback steckt schon in den Tools.“
31. 2-Wochen-Fenster: ausprobieren, echte Anträge, Hindernisse früh melden
32. Abschluss: 3 Mitnahmen wiederholen + nächster Termin + ms365.schule

Gestaltung: modern, ruhig, workshop-tauglich, hohe Lesbarkeit aus 5 Metern. Keine Cliparts. Jede Abschnittsfolie mit großer Ziffer links.
```

---

## 2) Abschnitts-Prompts (empfohlen – bessere Qualität)

### A – Eröffnung & Agenda

```text
Erstelle 8 Folien als Workshop-Eröffnung (Deutsch, Moderatoren-Stil, wenig Text):

1) Titelfolie
Überschrift: MS365 für die Schulverwaltung
Große Zeile: TAG 3
Unterzeile: Bauen · Mitnehmen · Entscheiden
Klein: #kurtrocks · ms365.schule

2) Zeitfolie
Nur groß: 09:30–16:15
Darunter: Freitag | Dritter Workshop-Tag

3) Branding
Groß zentriert: #kurtrocks
Darunter: edu.innovation.consulting

4) Einstieg online
Groß: ms365.schule
Kleiner: App-Zugang für heute bereithalten

5) Agenda „Unser Tag von 09:30 bis 16:15“
Sieben Zeilen mit Uhrzeit + Titel:
09:30 Ankommen & Versprechen
09:40 Login & Consent
10:10 Stammdaten
11:20 Mitnahme 1 Schulaktivitäten
13:00 Mitnahme 2 (Freistellung oder Matura)
13:45 Mitnahme 3 Kalender + Kursteams
14:05 Individuell · 15:35 Feedback & Close

6) Haltung
Titel: Heute kein Vortrag – wir bauen
Drei kurze Punkte:
- Ihr habt mitentwickelt
- Ich übermittle und bündele
- Mindestens 3 lauffähige Dinge mitnehmen

7) Rollen
Titel: Was jede Rolle heute spüren soll
3 Spalten:
Direktion – Genehmigen in Minuten, Überblick
Administration – Listen, Kalender, Alltag ohne Ticket
IT – Consent einmal, Stammdaten/Site, weniger Einzelanfragen

8) Versprechen
Titel: Eure 3 Mitnahmen
01 Schulaktivitäten live (Antrag → Genehmigung → Kalender)
02 Lehrerfreistellung oder Matura/sRDP
03 Kalender-Nutzen + kurze Kursteams-Demo (IT)
Kasten ERGEBNIS: Drei greifbare Ergebnisse – keine Folien-Demos
```

### B – Login & Consent

```text
Erstelle 4 Folien für den Block Login & Consent (09:40–10:10):

1) Abschnittsfolie
Große Ziffer: 1
Zeit: 09:40–10:10
Format: Plenum · jede Schule am Gerät
Titel: Rein in die App
ERGEBNIS: Jede Schule ist angemeldet und sieht das Portal

2) Login-Checkliste
Titel: Gemeinsam reinhelfen
01 URL öffnen
02 Organisationskonto der Schule
03 MFA durchkommen
04 Admin bestätigt Consent „für die Organisation“
05 Einmal ab- und wieder anmelden
06 Stockt >3 Min: Nachbarhilfe / Demo-Pfad, später nachziehen

3) Consent einfach
Titel: Warum die gelbe Microsoft-Seite kommt
Kurzpunkte:
- App speichert nichts beim Anbieter
- Arbeitet in eurem Microsoft-365-Tenant
- Braucht oft einmalig Administratorzustimmung
- Ohne Consent keine Listen/Prozesse im Tenant
Seitenkasten: „Das ist Microsoft-Sicherheitsmodell – kein Haken.“

4) Dealbreaker
Titel: Wer muss im Raum sein?
Groß: Global Admin oder Application Administrator
Zusatz: IT richtet ein · Fachperson erlebt den Flow
```

### C – Stammdaten

```text
Erstelle 3 Folien für Stammdaten (10:10–11:20):

1) Abschnittsfolie
Ziffer 2 · 10:10–11:20 · Schulteam + Begleitung
Titel: Stammdaten – eure Basis
ERGEBNIS: Klassen + Lehrkräfte (mit E-Mail) + Site gewählt

2) Minimum heute
Zwei Spalten:
Muss: Klassen-Codes | Lehrkräfte + E-Mail | Schuljahr | Site/Listen-Ort
Später: Schüler voll | WebUntis komplett | SharePoint-Sync
Hinweis: Keine Exports? 5–8 Klassen + 5 Lehrer manuell reichen für den ersten Ballon.

3) Co-Building
Titel: Wir füllen eure Basis
Zitat groß: „Nicht meine Demo-Daten – eure Schule.“
Punkte: WebUntis-Export / Vorlagen / Copy-Paste
Seitenkasten: Zeit lassen – das ist der eigentliche Zeitfresser und die Grundlage für alles Weitere.
```

### D – Mitnahme 1 Schulaktivitäten

```text
Erstelle 4 Folien + 1 Pausenfolie für Schulaktivitäten (11:20–12:15):

1) Abschnittsfolie
Ziffer 3 · 11:20–12:15 · Live im Tenant
Titel: Mitnahme 1 – Schulaktivitäten
ERGEBNIS: Ein vollständiger Antrag-Flow live

2) Prozesskette
Sechs Schritte horizontal oder als 01–06:
Melden → Prüfen → Entscheiden → Informieren → Kalender → In SharePoint speichern
Untertitel: Digitale Meldung und Genehmigung einer Schulveranstaltung/Exkursion

3) Arbeitsfokus
Titel: ACHTET AUF
Was überzeugt?
Was fehlt für unseren Schulalltag?
Welcher Sonderfall fehlt noch?
Unten: Danach zeigen – Eintrag liegt in SharePoint; Formularpfad ohne Dauer-App möglich

4) Scope & Strategie
Titel: Klar abgegrenzt
- Keine automatische Untis-Eintragung im MVP
- Wahrheit = SharePoint-Liste + Kalender/ICS
- App = schneller Einstieg und IT-Werkstatt
Pausenfolie danach: PAUSE | 12:15 Mittag
```

### E – Mitnahme 2 Freistellung / Matura

```text
Erstelle 4 Folien für den Nachmittag-Start (13:00–13:45):

1) Abschnittsfolie
Ziffer 4 · 13:00–13:45
Titel: Mitnahme 2 – zweiter lauffähiger Flow
ERGEBNIS: Schule hat einen zweiten Prozess zum Mitnehmen
Hinweis: Primär nach Schmerz wählen

2) Wahlfolie (zwei große Kacheln)
Links: Lehrerfreistellung – „Ich möchte frei haben“ – Direktion genehmigt – Kalender/iCal
Rechts: Matura / sRDP – Anmeldung – Lehrerblick – Auswertung – „nach Weihnachten“
Unten: HAK → Matura priorisieren | sonst Freistellung live, Matura als Kurzdemo

3) Lehrerfreistellung kompakt
01 Antrag Lehrkraft
02 Entscheidung Direktion
03 Übersicht / Kalender
Seitenkasten: Schüler-Freistellung = anderes Tool (1 Satz)

4) Matura/sRDP kompakt
01 Profil / Abschlussjahr
02 Testanmeldung
03 Rollen: Schüler · Lehrer · Admin
Seitenkasten: Keine Klausurplanung, keine zentrale srdp.at-Anbindung heute
```

### F – Mitnahme 3 Kalender + Kursteams

```text
Erstelle 3 Folien (13:45–14:05):

1) Abschnittsfolie
Ziffer 5 · 13:45–14:05 · Demo + kurzer Hands-on
Titel: Mitnahme 3 – Kalender-Nutzen & IT-Hebel
ERGEBNIS: Admin hat Kalender-Pfad · IT sieht den Hebel

2) Kalender
Titel: Was ihr sofort nutzen könnt
- Kalenderansicht im Prozess
- ICS-Download / Feed
- Sekretariat & Direktion behalten Überblick ohne Zettel
Satz: „Das arbeitet für euch weiter – auch wenn die Workshop-Session vorbei ist.“

3) Kursteams-Demo
Titel: 8–12 Minuten IT-Wow
- Kursteams anlegen/verwalten sichtbar machen
- Bewusst Demo, kein Deep-Dive
- Satz: „Neben Lehrer-Prozessen: so entlastet ihr die IT.“
Seitenkasten: Nicht Pflichtprogramm – Ausblick auf den Hebel
```

### G – Individuell

```text
Erstelle 2 Folien (14:05–15:35):

1) Abschnittsfolie
Ziffer 6 · 14:05–15:35 · Schulteam, Moderator wandert
Titel: Jetzt euer Fall
ERGEBNIS: Vertiefung oder schulischer Extra-Use-Case · 3er-Liste abgehakt

2) Arbeitsmodus
01 Eigenen Use Case öffnen / vertiefen
02 Fehlt noch Mitnahme 2 oder 3: nachziehen
03 Owner Fach + Owner IT notieren
04 Offene Frage und nächsten Schritt festhalten
Leitplanken: testbarer Mindestumfang · kein Parallelchaos · Close um 15:35 ist heilig
```

### H – Feedback & Close

```text
Erstelle 5 Folien für den Abschluss (15:35–16:15):

1) Abschnittsfolie
Ziffer 7 · 15:35–16:15 · Plenum
Titel: Feedback & nächste zwei Wochen
ERGEBNIS: Owner · Termin · 3 Mitnahmen bestätigt

2) Feedback-Würdigung
Titel: Ihr habt das mitgeprägt
Punkte:
- Feedback der letzten Wochen steckt schon in den Tools
- Heute wurde es in eurem Tenant konkret
- Weiteres Feedback steuert die nächsten zwei Wochen

3) 2-Wochen-Fenster
Titel: Bis zum Gespräch in ca. zwei Wochen
01 Die 3 Mitnahmen echt nutzen
02 Mit vereinbartem Prüffall testen
03 Hindernisse früh melden: Was fehlt, von wem, bis wann?
04 Kurz notieren: Was läuft / was hakt

4) Close ohne Druck
Titel: Ruhig entscheiden – klarer Pfad
Groß: Zwei Wochen Probe mit dem Werkzeug
Dann: Gespräch über Weiterarbeit (Schulalltag / Starter / Plus)
Daten bleiben im Schul-Tenant
Unten klein (optional, nicht dominant): Jahresbundle und Setup nur wenn gefragt – Fokus auf Termin & Owner

5) Schlussfolie
Groß: 3 Dinge mitnehmen. 2 Wochen arbeiten. Dann entscheiden.
Darunter: ms365.schule
Klein: #kurtrocks edu.innovation.consulting
```

---

## 3) Einzelne Feinschliff-Prompts (Copy-Paste pro Folie)

### Titelfolie nachschärfen

```text
Überarbeite die Titelfolie: maximale Klarheit, wenig Text. Oben klein „Workshop-Reihe“, Mitte „MS365 für die Schulverwaltung“, darunter sehr groß „TAG 3“, darunter „Bauen · Mitnehmen · Entscheiden“. Entferne Firmen-Blabla und Cliparts.
```

### Agenda als Zeitstrahl

```text
Gestalte die Agenda-Folie als vertikalen Zeitstrahl von 09:30 bis 16:15 mit genau 7 Stationen. Jede Station: Uhrzeit links, Kurztitel rechts. Hebe „15:35 Close“ optisch hervor. Keine Paragraphen.
```

### Consent-Folie entschärfen (nicht technisch wirken)

```text
Vereinfache die Consent-Folie auf Alltagssprache. Maximal 4 Bullet-Punkte. Keine API-Namen wie Sites.ReadWrite.All. Betone: Daten im Schul-Tenant, einmalige Admin-Zustimmung, ohne Consent kein Weiterkommen.
```

### Mitnahmen-Folie härter machen

```text
Mach die 3-Mitnahmen-Folie entscheidungsstärker: große Ziffern 1 2 3, je eine Zeile Ergebnis („lauffähig“, „im Tenant“, „mit Kalender“). Untertitel: „Keine Folien-Demos – drei Dinge zum Mitnehmen.“
```

### Close-Folie ohne Verkaufsdruck

```text
Entferne Preisangaben von der Abschlussfolie. Fokus nur: 3 Mitnahmen · 2 Wochen Probe · Owner · Termin. Verkauf/Preise bleiben im Gespräch, nicht auf der Folie.
```

### Pausenfolie

```text
Erstelle eine minimale Pausenfolie: nur das Wort PAUSE und die Uhrzeit, zentriert, viel Weißraum.
```

---

## 4) Optional: Moderatoren-Notizen (Copilot Speaker Notes)

```text
Ergänze auf allen Abschnittsfolien (Ziffer 1–7) Speaker Notes für den Moderator in Deutsch:
- 2–3 Sprechsätze
- Was live zu tun ist (Browser/App)
- Was bewusst NICHT gesagt wird (kein Vorstellen, kein 107-Kacheln-Katalog, kein Kaufdruck vor 15:35)
Halte die Notizen kurz und vortragbar.
```

### Speaker-Notes-Inhalte (zum Einfügen / Gegenprüfen)

| Folie | Notiz-Kern |
|-------|------------|
| Haltung | „Tag 3 – wir kennen uns. Heute bauen und mitnehmen.“ |
| Consent | „90 Sekunden Erklärung, dann helfen – nicht dozieren.“ |
| Stammdaten | „Zeit lassen. Minimum reicht. Perfektion tötet den Ballon.“ |
| Schulaktivitäten | „Ein kompletter Flow schlägt drei halbe Demos.“ |
| Mitnahme 2 | „Nach Schmerz wählen: Freistellung oder Matura.“ |
| Kursteams | „Demo-Hebel für IT – nicht Pflichtprogramm.“ |
| Individuell | „Ich wandere. Close um 15:35 nicht opfern.“ |
| Close | „Feedback würdigen → 3 Mitnahmen → 2 Wochen → Termin.“ |

---

## 5) Folien-Landkarte (Checkliste zum Abhaken)

| # | Folie | Im Deck? |
|---|--------|----------|
| 1 | Titel TAG 3 | ☐ |
| 2 | 09:30–16:15 | ☐ |
| 3 | #kurtrocks | ☐ |
| 4 | ms365.schule | ☐ |
| 5 | Agenda Zeitstrahl | ☐ |
| 6 | Haltung / Co-Building | ☐ |
| 7 | Rollen-Wow | ☐ |
| 8 | 3 Mitnahmen | ☐ |
| 9 | Abschn. Login | ☐ |
| 10 | Login-Checkliste | ☐ |
| 11 | Consent einfach | ☐ |
| 12 | Dealbreaker Admin | ☐ |
| 13 | Abschn. Stammdaten | ☐ |
| 14 | Minimum heute | ☐ |
| 15 | Co-Building Basis | ☐ |
| 16 | Abschn. Schulaktivitäten | ☐ |
| 17 | Prozesskette | ☐ |
| 18 | Achtet auf | ☐ |
| 19 | Scope/Strategie | ☐ |
| 20 | Pause Mittag | ☐ |
| 21 | Abschn. Mitnahme 2 | ☐ |
| 22 | Wahl Freistellung/Matura | ☐ |
| 23 | Freistellung kompakt | ☐ |
| 24 | Matura kompakt | ☐ |
| 25 | Abschn. Mitnahme 3 | ☐ |
| 26 | Kalender | ☐ |
| 27 | Kursteams-Demo | ☐ |
| 28 | Abschn. Individuell | ☐ |
| 29 | Arbeitsmodus | ☐ |
| 30 | Abschn. Close | ☐ |
| 31 | Feedback-Würdigung | ☐ |
| 32 | 2-Wochen-Fenster | ☐ |
| 33 | Close ohne Druck | ☐ |
| 34 | Schluss ms365.schule | ☐ |

---

## 6) Was bewusst **nicht** auf die Folien

- Lange Vorstellung / Lebenslauf / Firmenhistorie  
- 107 Tool-Kacheln / Admin-Katalog  
- BIP / Bildungsportal / Intune / Noten  
- Preise dominant (höchstens mündlich im Close)  
- Versprechen Auto-Untis / Auto-Sokrates  
- Echtdaten-Produktivsetzung „fertig bis heute Abend“

---

## Verwandt

- Playbook: [`verkauf-freitag-workshop.md`](./verkauf-freitag-workshop.md)  
- Vorbild-PPTX: Block B vom 06.11.2026 (Termin 2)
