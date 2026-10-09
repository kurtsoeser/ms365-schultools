# Freistellungen – Power Automate Flow v2 (Bauanleitung)

Ziel: Genehmigung in **Microsoft Approvals** (Teams/Handy), Regeln **1 Tag / mehrtägig / KV = Direktion**, Audit-Felder in SharePoint für den Planer.

Voraussetzung: Liste „Freistellungen“ mit Spalten aus dem **Freistellungen-Setup** (inkl. Audit-Spalten `GenehmigtVonKV`, …).

## Konfiguration (später ohne Flow-Designer änderbar)

Empfohlen in der Solution als **Umgebungsvariablen** (Text):

| Variable | Beispiel | Verwendung |
|----------|----------|------------|
| `FR_EmailDirektion` | direktion@schule.at | Sequential Schritt 2, Vergleich KV=Direktion |
| `FR_EmailSonder` | vertretung@schule.at | Eine Genehmigung wenn KV = Direktion |
| `FR_EmailPostfach` | automate@schule.at | Shared Mailbox Senden |
| `FR_SiteUrl` | https://…/sites/Administration | SharePoint-Trigger/Aktionen |
| `FR_ListId` | GUID | Liste |

Im ZIP-Setup werden dieselben Werte heute per String-Ersetzung eingetragen. Beim Bau in PA v2 können Sie feste E-Mails eintragen und später auf Variablen umstellen.

## Schritt 0 – Trigger

- **Wenn ein Element erstellt wird** (SharePoint) – Site + Liste `Freistellungen`
- Optional: Bedingung `Status` = `Ausstehend` (falls Status beim Anlegen gesetzt wird)

## Schritt 1 – Tagesanzahl (inklusive)

**Compose** `TageAnzahl`:

```text
add(div(sub(ticks(formatDateTime(triggerBody()?['Ende'], 'yyyy-MM-dd')),
           ticks(formatDateTime(triggerBody()?['Beginn'], 'yyyy-MM-dd'))),
        86400000000000),
    1)
```

(Oder: zwei **Convert time zone** + Differenz – wichtig ist: **1 Tag** = gleicher Beginn/Ende, **≥ 2** = mehrtägig.)

**Compose** `KvEmail`:

```text
toLower(triggerBody()?['Klassenvorstand/Email'])
```

## Schritt 2 – Hauptverzweigung KV = Direktion?

**Bedingung**:

```text
@equals(outputs('KvEmail'), toLower(parameters('FR_EmailDirektion')))
```

(Ohne Umgebungsvariable: feste Direktions-E-Mail aus dem Setup.)

### Zweig Ja – nur Sondergenehmigung

1. **Genehmigung starten und auf Antwort warten** – Typ **Basic**, Zugewiesen an `FR_EmailSonder`
2. Details wie bisher (Name, Klasse, Beginn, Ende, Kategorie, Link)
3. **Bedingung** Outcome = Approve
   - **Ja:** Element aktualisieren: `Status` = Genehmigt; Audit: `GenehmigtVonKV` = Antwort Genehmiger (siehe unten); `GenehmigtAmKV` = `utcNow()` (date only); `Bemerkungen` = Kommentar
   - **Nein:** `Status` = Abgelehnt; `AbgelehntVon`, `AbgelehntAm`
4. E-Mail aus freigegebenem Postfach an `Author/Email`
5. **Beenden** (Succeeded) – verhindert Doppelpfad

### Zweig Nein – 1 Tag vs. mehrtägig

**Bedingung** innen:

```text
@less(outputs('TageAnzahl'), 2)
```

#### 1 Tag – nur Klassenvorstand

- Approvals **Basic**, `assignedTo` = `@triggerBody()?['Klassenvorstand/Email']`
- Nach Antwort: Status + **nur** `GenehmigtVonKV` / `GenehmigtAmKV` (oder Ablehnung)

#### ≥ 2 Tage – Sequential KV → Direktion

- Approvals **Sequential**, Schritte:
  1. `@triggerBody()?['Klassenvorstand/Email']`
  2. `FR_EmailDirektion`
- **Apply to each** über `outputs('Genehmigung')?['body/responses']` → Kommentare in Array
- Erfolg: `contains(outputs('Genehmigung')?['body/outcome'],'Approve')` **und** Outcome nicht `Reject` (oder wie bisher `Approve, Approve` testen)
- Element aktualisieren:
  - `GenehmigtVonKV` = `@{first(body('Genehmigung')?['responses'])?['responder/displayName']}` (Schritt 1)
  - `GenehmigtAmKV` = Datum Schritt 1
  - `GenehmigtVonDirektion` / `GenehmigtAmDirektion` = Schritt 2
  - `Bemerkungen` = `join(variables('Kommentare'), …)`

## Audit-Felder aus Approval (Beispielausdrücke)

Nach **Basic**:

- Name: `@{outputs('Genehmigung_KV')?['body/responses'][0]['responder/displayName']}`
- E-Mail: `@{outputs('Genehmigung_KV')?['body/responses'][0]['responder/email']}`
- Kombiniert in `GenehmigtVonKV`: `@{concat(..., ' <', ..., '>')}`

Datum (date only): `@{formatDateTime(utcNow(), 'yyyy-MM-dd')}` oder `responseDate` der Approval.

## Vereinfachungen gegenüber v1

- **Title** der Liste nicht umbenennen (optional) – Planer liest `Title` + Felder
- Gemeinsame Unter-Flow-Aktionen: „Status setzen + Mail“ als **Scope** duplizieren vermeiden (ein Scope pro Ergebnis)
- Kein unnötiges `GenehmigungsArray` in der Schleife

## Testplan (HAK / Testliste)

| # | Fall | Erwartung |
|---|------|-----------|
| 1 | 1 Tag, KV ≠ Direktion | 1 Approval an KV |
| 2 | 3 Tage, KV ≠ Direktion | Sequential KV + Direktion |
| 3 | KV-E-Mail = Direktion | 1 Approval an Sonderperson |
| 4 | Ablehnung KV | Status Abgelehnt, `AbgelehntVon` gesetzt |

## Export ins Projekt

1. **Meine Flows** → Flow → **Export** → Package (Legacy)
2. ZIP in `assets/power-automate/freistellung/` ersetzen (Ordner-GUID = `FLOW_ASSET_ID` in `freistellung-setup.js`, aktuell v3: `6c60dd7e-ab68-4cc8-949e-d689badc0993`)
3. Platzhalter in `freistellung-setup.js` `SOURCE` an die Export-Werte anpassen
4. Setup-Tool testen: Paket laden, in Test-Tenant importieren

## Nach dem Export

- Konzept-Seite und Planer zeigen Audit-Spalten
- Schulen: **Liste anlegen / prüfen** im Setup → fehlende Spalten werden ergänzt
