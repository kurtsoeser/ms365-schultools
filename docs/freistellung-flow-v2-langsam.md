# Freistellungen Flow v2 – langsam bauen (Modul für Modul)

Arbeite **nur ein Modul**, speichere, teste wenn angegeben, dann weiter.  
Designer: [make.powerautomate.com](https://make.powerautomate.com) → **Erstellen** → **Automatisierter Cloud-Flow**.

E-Mails vorher notieren:

- Direktion: `________________________`
- Sonder (wenn KV = Direktion): `________________________`
- Freigegebenes Postfach: `________________________`

---

## Modul 1 – Flow anlegen, Trigger, Vorbereitung

**Ziel heute:** Flow existiert, reagiert auf neue Listeneinträge, `KvEmail` und `TageAnzahl` stehen bereit. **Noch keine Genehmigung.**

1. Name: `Freistellungen - Genehmigungsprozess v2`
2. Trigger: **Wenn ein Element erstellt wird** (SharePoint)  
   - Site: eure Administrations-Site  
   - Listenname: **Freistellungen**
3. **+ Neuer Schritt** → **Variable initialisieren**  
   - Name: `KommentareGenehmigung`  
   - Typ: **Array**
4. **+ Neuer Schritt** → **Erstellen** (Compose)  
   - Eingabe, Umbenennen in: `KvEmail`  
   - Ausdruck (dynamischer Inhalt + Expression):

```text
toLower(triggerBody()?['Klassenvorstand/Email'])
```

   Falls das Feld im Trigger anders heißt: im Dynamic content **Klassenvorstand Email** wählen und `toLower()` drumherum.

5. **+ Neuer Schritt** → **Erstellen** → Name: `TageAnzahl`  

```text
add(div(sub(ticks(formatDateTime(triggerBody()?['Ende'], 'yyyy-MM-dd')), ticks(formatDateTime(triggerBody()?['Beginn'], 'yyyy-MM-dd'))), 86400000000000), 1)
```

6. **Speichern**. Test: in SharePoint **ein Testelement** anlegen (Beginn = Ende = ein Datum). Flow-Lauf öffnen → Outputs von `KvEmail` und `TageAnzahl` prüfen (1 Tag → `TageAnzahl` = 1).

**Fertig Modul 1** → weiter mit Modul 2.

---

## Modul 2 – Sonderfall KV = Direktion

**Ziel:** Wenn KV-E-Mail = Direktion → **eine** Approval an Sonderperson, Status + Mail, **Beenden**.

1. Unter `TageAnzahl`: **+ Neuer Schritt** → **Bedingung** → Umbenennen: `KvIstDirektion`
2. Bedingung (Ausdruck):

```text
@equals(outputs('KvEmail'), 'DIREKTION@schule.at')
```

   (Kleinbuchstaben, echte Direktions-Mail einsetzen – gleich wie in `KvEmail` verglichen.)

3. **Wenn ja** (linker Zweig):
   - **Genehmigung starten und auf Antwort warten** (Approvals)  
     - Genehmigungstyp: **Benutzerdefinierte Genehmigung – nur eine Antwort** (Basic)  
     - Titel: `Antrag Freistellung`  
     - Zugewiesen an: Sonder-E-Mail  
     - Details: Autor, Klasse, Beginn, Ende, Kategorie, Beschreibung (Dynamic content)  
     - Link: Link zum Element  
     - Anforderer: Author Email  
   - **Bedingung** innen: Outcome = **Approve** (Dynamic content der Approval-Aktion)
     - **Ja:** SharePoint **Element aktualisieren** – ID vom Trigger, Status = `Genehmigt`, `GenehmigtVonKV` = Name/E-Mail aus Approval, `GenehmigtAmKV` = heute (date only), optional Bemerkungen = Kommentar  
     - **Nein:** Status = `Abgelehnt`, `AbgelehntVon`, `AbgelehntAm`  
   - **E-Mail aus freigegebenem Postfach senden** an Author Email (Betreff GENEHMIGT/ABGELEHNT)  
   - **Beenden** → Status **Succeeded**

4. **Wenn nein** (rechter Zweig): vorerst **leer lassen**.

5. Speichern. Test nur wenn KV testweise = Direktion (selten) – sonst Modul 3 zuerst testen.

**Fertig Modul 2** → Modul 3.

---

## Modul 3 – Ein Tag (nur KV)

Im **Nein**-Zweig von `KvIstDirektion`:

1. **Bedingung** `NurEinTag`: Ausdruck

```text
@less(outputs('TageAnzahl'), 2)
```

2. **Wenn ja** (1 Tag):
   - Approvals **Basic** → Zugewiesen an **Klassenvorstand Email** (Trigger)  
   - Wie Modul 2: Approve/Reject → SharePoint + Mail (nur KV-Audit-Felder bei Approve)  
   - **Kein Beenden** nötig (Ende des Zweigs)

3. Speichern. Test: 1 Tag, KV ≠ Direktion → genau **eine** Approval an KV.

**Fertig Modul 3** → Modul 4.

---

## Modul 4 – Mehrtägig (Sequential)

Im **Nein**-Zweig von `NurEinTag` (also ≥ 2 Tage):

1. Approvals → Typ **Sequentiell**  
   - Schritt 1: Klassenvorstand Email  
   - Schritt 2: Direktion-E-Mail (fest)
2. **Apply to each** über `responses` der Approval → Kommentare in `KommentareGenehmigung` anfügen (optional, für Bemerkungen)
3. Bedingung: Gesamt-Outcome genehmigt (z. B. enthält Approve und nicht Reject – oder `Approve, Approve` testen wie im alten Flow)
4. **Element aktualisieren:** Status, `GenehmigtVonKV` / `Am` aus Antwort 1, `GenehmigtVonDirektion` / `Am` aus Antwort 2, `Bemerkungen` = join Kommentare
5. Mail an Autor

Test: Beginn–Ende = 3 Tage, zwei Approvals nacheinander.

**Fertig Modul 4** → Modul 5.

---

## Modul 5 – Aufräumen

- Überall: **Title** nicht mehr setzen (falls Copilot/Altlast)  
- Status nur `Genehmigt` / `Abgelehnt` (Schreibweise wie in der Liste)  
- Alle 4 Testfälle (Tabelle in `freistellung-flow-v2-anleitung.html`)  
- Flow **einschalten**

---

## Modul 6 – Export

Package (Legacy) → Projekt `assets/power-automate/freistellung/` → Setup `SOURCE` anpassen.

---

## Testdaten (ms365.schule)

Datei: `docs/demo-data/freistellungen-test-ms365-schule.json`  
Planer → **Testdaten JSON** → Datei wählen → optional SharePoint (Flow vorher **aus** für neue Ausstehend-Zeilen).

## Planer-Rollen (Entra)

Im **Freistellungen-Setup** (Schritt Liste): Entra-Gruppen für Schüler, Klassenvorstände und Direktion eintragen.  
Der Planer nutzt `Group.Read.All` und ordnet nach Anmeldung die Ansicht zu (Filter „Meine Anträge“ / KV / gesamte Schule).  
Zusätzlich: KV über **KV-E-Mail der Klasse** in den Stammdaten; Direktion über **Direktion-E-Mail** im Setup.  
IT-Vorschau mit Rollen-Umschalter: `?demoRole=1` an die Planer-URL (nicht für Schüler im Alltag).

---

## Wenn du hängen bleibst

Schreib z. B. „Modul 2, Schritt Bedingung“ + Screenshot oder Fehlermeldung – dann nur diesen einen Schritt klären.
