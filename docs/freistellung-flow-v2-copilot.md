# Freistellungen Flow v2 – Prompts für Copilot in Power Automate

## Geht das „einfach so“?

**Teilweise – ja.** Copilot in Power Automate (Designer → Copilot / „Describe it to design it“) kann einen **guten ersten Entwurf** liefern: Trigger, Bedingungen, Genehmigungs-Aktionen, SharePoint-Update, E-Mail.

**Vollautomatisch und fehlerfrei – eher nein**, wenn:

- **Sequentielle** Microsoft Approvals mit **zwei Genehmigern** nötig sind,
- **drei Verzweigungen** (KV = Direktion / 1 Tag / mehrtägig) kombiniert werden,
- **Audit-Spalten** exakt benannt werden müssen,
- **Personenfeld** `Klassenvorstand/Email` im Trigger korrekt referenziert wird.

**Empfehlung:** Copilot in **4–5 kurzen Prompts** nacheinander nutzen (oder einen Master-Prompt + **Nachprompts** zum Korrigieren). Danach **15–30 Minuten manuell prüfen** (Checkliste unten).

---

## Vor Copilot vorbereiten

1. SharePoint-Liste **Freistellungen** existiert (MS365-Schul-Tools → Setup → „Liste anlegen / prüfen“).
2. Du bist in der **richtigen Power-Platform-Umgebung** (Schule, nicht Privat).
3. Connections: **SharePoint**, **Approvals**, **Office 365 Outlook** (für freigegebenes Postfach) – Copilot fragt ggf. danach.
4. Diese Platzhalter-E-Mails bereithalten (später anpassbar):
   - `DIREKTION_EMAIL` = z. B. direktion@schule.at  
   - `SONDER_EMAIL` = z. B. vertretung@schule.at (wenn KV = Direktion)  
   - `POSTFACH_EMAIL` = z. B. automate@schule.at  

5. Neuen **Cloud-Flow** anlegen (automatisch) oder Copilot auf leerem Flow starten.

---

## Strategie A: Vier Prompts (empfohlen)

Jeden Prompt **einzeln** in Copilot eingeben, warten bis der Flow aktualisiert ist, kurz prüfen, dann den nächsten.

### Prompt 1 – Grundgerüst + Trigger

```text
Erstelle einen automatisierten Cloud-Flow für eine Schule (Österreich, Deutsch).

Trigger: Wenn ein Element in einer SharePoint-Liste erstellt wird.
Liste: Freistellungen auf meiner SharePoint-Website (ich wähle Site und Liste im Trigger selbst aus).

Die Liste hat diese Spalten (interne Namen):
Beginn, Ende, Status (Ausstehend | Genehmigt | Abgelehnt), Klasse, Klassenvorstand (Person),
Kategorie, Beschreibung, Bemerkungen,
GenehmigtVonKV, GenehmigtAmKV, GenehmigtVonDirektion, GenehmigtAmDirektion, AbgelehntVon, AbgelehntAm.

Nach dem Trigger:
- Initialisiere eine Array-Variable "KommentareGenehmigung" (leer).
- Compose-Aktion "KvEmail": toLowercase der E-Mail des Klassenvorstands aus dem Trigger (SharePoint-Feld Klassenvorstand Email).
- Compose-Aktion "TageAnzahl": inklusive Anzahl Kalendertage zwischen Spalte Beginn und Spalte Ende (beide zählen mit; gleicher Tag = 1 Tag). Nutze Ausdruck mit ticks/formatDateTime.

Noch keine Genehmigung in diesem Schritt – nur Trigger und diese Vorbereitung.
```

### Prompt 2 – Verzweigung KV = Direktion + Sonderfall

```text
Erweitere den Flow nach den Compose-Aktionen:

Bedingung "KvIstDirektion":
Wenn KvEmail gleich DIREKTION_EMAIL ist (hardcode vorerst: DIREKTION_EMAIL = direktion@schule.at, ich ersetze später).

Wenn ja (True):
- Aktion "Genehmigung starten und auf Antwort warten" (Microsoft Approvals), Typ Basic,
  Titel "Antrag Freistellung", zugewiesen an SONDER_EMAIL (vertretung@schule.at),
  Details mit Name des Autors, Klasse, Beginn, Ende, Kategorie, Beschreibung, Link zum Element.
  Requestor = E-Mail des Autors aus dem Trigger.
- Danach Bedingung ob Genehmigung outcome Approve ist.
  - Wenn Approve: SharePoint-Element aktualisieren – Status = Genehmigt,
    GenehmigtVonKV und GenehmigtAmKV aus der Approval-Antwort (Anzeigename, E-Mail, Datum),
    Bemerkungen = Kommentar aus Approval.
  - Wenn nicht Approve: Status = Abgelehnt, AbgelehntVon und AbgelehntAm setzen.
- E-Mail aus freigegebenem Postfach (Office 365) an Autor-E-Mail, Betreff enthält GENEHMIGT oder ABGELEHNT.
- Aktion Beenden (Succeeded) am Ende des True-Zweigs, damit der False-Zweig nicht auch läuft.

Wenn nein (False): noch leer lassen, kommt im nächsten Prompt.
```

*(Vor dem Einfügen: `direktion@schule.at` und `vertretung@schule.at` durch echte Adressen ersetzen.)*

### Prompt 3 – 1 Tag vs. mehrtägig im False-Zweig

```text
Im False-Zweig von "KvIstDirektion" (wenn Klassenvorstand nicht die Direktion ist):

Neue Bedingung "Mehrtaegig":
Wenn TageAnzahl größer oder gleich 2 ist.

Wenn nein (nur 1 Tag):
- Microsoft Approvals Basic, zugewiesen an Klassenvorstand-E-Mail aus dem Trigger (Klassenvorstand Email).
- Nach Antwort: bei Approve Status Genehmigt, nur GenehmigtVonKV und GenehmigtAmKV füllen;
  bei Ablehnung AbgelehntVon/AbgelehntAm.
- E-Mail aus freigegebenem Postfach POSTFACH_EMAIL an den Autor.

Wenn ja (2 oder mehr Tage):
- Microsoft Approvals Sequential mit zwei Schritten:
  Schritt 1: Klassenvorstand-E-Mail aus Trigger,
  Schritt 2: DIREKTION_EMAIL (direktion@schule.at).
- Nach Antwort: bei vollständiger Genehmigung Status Genehmigt,
  GenehmigtVonKV aus erster Antwort, GenehmigtAmKV Datum erste Antwort,
  GenehmigtVonDirektion und GenehmigtAmDirektion aus zweiter Antwort,
  alle Approval-Kommentare in Bemerkungen zusammenfassen (Apply to each über responses).
- Bei Ablehnung: Status Abgelehnt, AbgelehntVon/AbgelehntAm vom ablehnenden Schritt.
- E-Mail aus freigegebenem Postfach an Autor.

Verwende für POSTFACH_EMAIL vorerst automate@schule.at.
```

### Prompt 4 – Qualität & keine Titel-Änderung

```text
Überprüfe den gesamten Flow und korrigiere:

- Ändere NICHT das SharePoint-Feld Title beim Aktualisieren (nur Status und Audit-Felder).
- Status-Werte exakt: Genehmigt, Abgelehnt, Ausstehend (Großschreibung wie angegeben).
- Alle SharePoint-Updates müssen dieselbe Liste und dieselbe Item-ID wie der Trigger verwenden.
- Keine doppelten E-Mails an den Autor pro Durchlauf.
- Im Sonderfall (KvIstDirektion True) muss nach Beenden kein weiterer Genehmigungspfad laufen.

Fasse mir am Ende kurz auf Deutsch zusammen, welche Bedingungen es gibt.
```

### Prompt 5 (optional) – Umgebungsvariablen

```text
Ersetze die hardcodierten E-Mail-Adressen für Direktion, Sondergenehmigung und Postfach
durch Power Automate Umgebungsvariablen vom Typ Text:
FR_EmailDirektion, FR_EmailSonder, FR_EmailPostfach.
Falls Variablen noch nicht existieren, erkläre mir wo ich sie anlegen muss, aber behalte Fallback auf die bisherigen Adressen in Kommentaren.
```

*(Copilot schafft das oft nicht zuverlässig – Variablen ggf. manuell im Admin Center anlegen.)*

---

## Strategie B: Ein Master-Prompt (einmal einfügen)

Wenn Copilot einen langen Prompt verkürzt oder Teile vergisst, auf **Strategie A** wechseln.

```text
Baue einen Cloud-Flow (Deutsch, Schule Österreich):

TRIGGER: SharePoint – wenn ein Element erstellt wird, Liste "Freistellungen".

LISTENFELDER: Beginn, Ende, Status (Choice: Ausstehend, Genehmigt, Abgelehnt), Klasse, Klassenvorstand (Person mit Email), Kategorie, Beschreibung, Bemerkungen, GenehmigtVonKV, GenehmigtAmKV, GenehmigtVonDirektion, GenehmigtAmDirektion, AbgelehntVon, AbgelehntAm.

VORBEREITUNG:
- Variable Array KommentareGenehmigung
- Compose KvEmail = toLower(Klassenvorstand Email vom Trigger)
- Compose TageAnzahl = inklusive Kalendertage zwischen Beginn und Ende (1 wenn gleicher Tag)

LOGIK:
1) Wenn KvEmail equals direktion@schule.at (Sonderfall KV ist Direktion):
   - Approvals Basic an vertretung@schule.at
   - Approve → Status Genehmigt, Audit-Felder KV aus Response, Mail an Autor aus automate@schule.at
   - Reject → Status Abgelehnt, AbgelehntVon/Am, Mail
   - Terminate Succeeded

2) Sonst wenn TageAnzahl < 2 (ein Tag):
   - Approvals Basic an Klassenvorstand Email
   - Approve/Reject wie oben, Audit nur KV-Felder bei Approve

3) Sonst (2+ Tage):
   - Approvals Sequential: Schritt1 Klassenvorstand Email, Schritt2 direktion@schule.at
   - Approve beide → Status Genehmigt, beide Audit-Paare füllen, Kommentare in Bemerkungen
   - Reject → Abgelehnt mit AbgelehntVon/Am

REGELN:
- Title nicht ändern
- Microsoft Approvals Connector (Teams)
- E-Mail: Send email from a shared mailbox (Office 365)
- Eine E-Mail pro Durchlauf an den Antragsteller (Author Email)
```

*(E-Mail-Adressen vor dem Einfügen anpassen.)*

---

## Nachprompts, wenn Copilot falsch liegt

| Problem | Prompt an Copilot |
|--------|-------------------|
| Sequential fehlt, nur Basic | „Im Zweig Mehrtaegig: ersetze Basic durch Approvals Sequential mit zwei Schritten: zuerst Klassenvorstand Email, dann direktion@schule.at.“ |
| Beide Zweige laufen | „Nach dem Sonderfall-Zweig (KvIstDirektion = ja) füge Aktion Beenden (Succeeded) hinzu, damit der False-Zweig nicht ausgeführt wird.“ |
| Falsches Outcome | „Prüfe Sequential: Erfolg nur wenn outcome Approve enthält und nicht Reject; bei zwei Schritten beide Approvals auswerten.“ |
| Title wird geändert | „Entferne bei allen SharePoint-Updates die Zuweisung zum Feld Title.“ |
| Keine Audit-Spalten | „Bei jedem Approve: schreibe GenehmigtVonKV als 'Name <email>' aus responses[0]; GenehmigtAmKV als yyyy-MM-dd; bei Sequential zusätzlich GenehmigtVonDirektion/Am aus responses[1].“ |
| Trigger pollt nur | „Stelle den SharePoint-Trigger auf 'When an item is created' (nicht nur geändert), falls verfügbar.“ |

---

## Checkliste nach Copilot (manuell, 10 Punkte)

- [ ] Trigger: richtige **Site** und Liste **Freistellungen**
- [ ] `KvEmail` nutzt **Klassenvorstand/Email** (Dynamic content)
- [ ] Sonderfall: **eine** Approval, dann **Beenden**
- [ ] 1 Tag: nur **eine** Approval an KV
- [ ] ≥2 Tage: **Sequential** mit 2 Assignees
- [ ] Status nur: `Genehmigt` / `Abgelehnt` (Choice exakt)
- [ ] **Title** wird nicht überschrieben
- [ ] Audit-Spalten werden bei Erfolg/Ablehnung gesetzt
- [ ] Mail: **Shared mailbox** + richtige Absender-Adresse
- [ ] Test mit 4 Fällen (siehe `freistellung-flow-v2-anleitung.html`)

---

## Export fürs MS365-Schul-Tools-Projekt

Wie in der Bauanleitung: **Export Package (Legacy)** → ZIP nach `assets/power-automate/freistellung/` → `SOURCE` in `freistellung-setup.js` anpassen.

---

## Kurzantwort für dich

**Ja, Copilot kann dir den Großteil der Arbeit abnehmen** – am besten mit **Prompt 1–4 nacheinander**. **Nein, blind vertrauen** solltest du nicht; Genehmigungslogik und Audit-Felder sind der Punkt, an dem Copilot oft nachbessern muss. Die Nachprompt-Tabelle und Checkliste sind dafür gedacht.
