# MS365-Schulverwaltung

Einfache **reine Browser-App** (ohne Server): Exportdaten (z. B. WebUntis CSV/Excel) werden aufbereitet, damit Sie daraus **Kursteams**, **Jahrgangsgruppen** und **ARGEs** per **PowerShell** anlegen können.

## Datenschutz

Es werden **keine Daten** an einen Server gesendet. Verarbeitung erfolgt **lokal im Browser**. Optional können Sie einen Zwischenstand im **lokalen Speicher** dieses Browsers sichern.

## Nutzung

1. Repository klonen oder die Dateien herunterladen.
2. `webuntis-teams-creator.html` im Browser öffnen **oder** ein statisches Hosting nutzen (z. B. GitHub Pages).
3. Oben zwischen **Kursteams**, **Jahrgangsgruppen** und **ARGEs** wählen.

### Modus A: Kursteams

1. Fächerliste exportieren und durch die Schritte in der App führen.
2. Entweder `neueteams.csv` + kurzes Skript aus der Anleitung **oder** das Paket **Kursteam-Anlage.ps1** und **Kursteam-Anlage.cmd** herunterladen, beide in denselben Ordner legen und die **CMD** per Doppelklick starten (Anmeldung interaktiv für MFA oder optional per `Get-Credential`).

### Modus B: Jahrgangsgruppen (Microsoft 365-Gruppen)

1. Domain und Präfix (z. B. `jg`) einstellen; Klassenzeilen im Format `1AK;2030` (Klasse;Abschlussjahr) eintragen.
2. Besitzer-UPNs pro Klasse eintragen.
3. Das generierte **Microsoft Graph PowerShell**-Skript kopieren **oder** **Jahrgangsgruppen-Anlage.ps1** + **.cmd** herunterladen und die CMD starten (interaktive Graph-Anmeldung, MFA möglich).

Dokumentation: [New-MgGroup](https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup), [Gruppe erstellen (Graph)](https://learn.microsoft.com/graph/api/group-post-groups).

### Modus C: ARGEs (Microsoft 365-Gruppen)

1. Domain angeben und optional Präfix/Schreibweise für den Mail-Nickname festlegen.
2. ARGE-Zeilen eintragen (entweder nur Anzeigename oder `Anzeigename;MailNickname`, z. B. `ARGE BB;ARGEBB`).
3. Besitzer-UPNs pro ARGE eintragen.
4. Script kopieren oder **ARGE-Gruppen-Anlage.ps1** + **.cmd** herunterladen und die CMD starten (interaktive Graph-Anmeldung, MFA möglich).

### Erwartete Spalten (flexibel)

Die App erkennt u. a.: **Klasse(n)**, **Fach**, **Lehrer**, **Schülergruppe** (je nach Export unterschiedlich benannt).

### Dateien

| Datei | Beschreibung |
|--------|----------------|
| `webuntis-teams-creator.html` | Hauptseite |
| `app.js` | Logik Kursteams (CSV-Export, Filter, lokaler Speicher, …) |
| `jahrgang.js` | Assistent Jahrgangsgruppen (Microsoft Graph PowerShell) |
| `arge.js` | Assistent ARGEs (Microsoft Graph PowerShell) |
| `index.html` | Optional: Weiterleitung zur Hauptseite (für GitHub Pages-Start-URL) |

## GitHub Pages

Repository auf **Pages** schalten (Branch `main`, Ordner `/`). Mit dem Repo-Namen **`ms365-schultools`** ist die App typischerweise unter:

- **Startseite:** [https://kurtsoeser.github.io/ms365-schultools/](https://kurtsoeser.github.io/ms365-schultools/)
- **Haupt-App:** [https://kurtsoeser.github.io/ms365-schultools/webuntis-teams-creator.html](https://kurtsoeser.github.io/ms365-schultools/webuntis-teams-creator.html)

Klonen per Git (nach tatsächlichem Repo-Namen auf GitHub):

```bash
git clone https://github.com/kurtsoeser/ms365-schultools.git
```

### Lokales Git nach Umbenennung des Repos auf GitHub

Wenn Sie das Repository auf GitHub umbenannt haben, passen Sie die **Remote-URL** in Ihrem geklonten Ordner an (einmalig):

```bash
git remote set-url origin https://github.com/kurtsoeser/ms365-schultools.git
git remote -v
```

**Hinweis:** Bei kostenlosem GitHub ist Pages für **private** Repos oft nicht verfügbar; öffentliches Repo oder GitHub Pro nötig.

## Lizenz

Keine Lizenz gesetzt – ergänzen Sie bei Bedarf eine `LICENSE` im Repository.
