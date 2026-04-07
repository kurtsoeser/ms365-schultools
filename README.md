# WebUntis → Microsoft Teams (Kursteam-Creator)

Einfache **reine Browser-App** (ohne Server): WebUntis-Exporte (CSV/Excel) werden aufbereitet, damit Sie daraus **Microsoft-Kursteams** per **PowerShell** (`New-Team -Template "EDU_Class"`) anlegen können.

## Datenschutz

Es werden **keine Daten** an einen Server gesendet. Verarbeitung erfolgt **lokal im Browser**. Optional können Sie einen Zwischenstand im **lokalen Speicher** dieses Browsers sichern.

## Nutzung

1. Repository klonen oder die Dateien herunterladen.
2. `webuntis-teams-creator.html` im Browser öffnen **oder** ein statisches Hosting nutzen (z. B. GitHub Pages).
3. Oben zwischen **Kursteams (WebUntis)** und **Jahrgangsgruppen (M365-Gruppen)** wählen.

### Modus A: Kursteams (WebUntis)

1. WebUntis-Fächerliste exportieren und durch die Schritte in der App führen.
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

Repository auf **Pages** schalten (Branch `main`, Ordner `/`). Die App ist dann unter  
`https://<user>.github.io/<repo>/` erreichbar – mit `index.html` als Einstieg oder direkt  
`…/webuntis-teams-creator.html`.

**Hinweis:** Bei kostenlosem GitHub ist Pages für **private** Repos oft nicht verfügbar; öffentliches Repo oder GitHub Pro nötig.

## Lizenz

Keine Lizenz gesetzt – ergänzen Sie bei Bedarf eine `LICENSE` im Repository.
