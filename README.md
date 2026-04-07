# MS365-Schulverwaltung

Einfache **reine Browser-App** (ohne Server): Exportdaten (z. B. WebUntis CSV/Excel) werden aufbereitet, damit Sie daraus **Kursteams**, **Jahrgangsgruppen** und **ARGEs** per **PowerShell** anlegen können.

## Datenschutz

Es werden **keine Daten** an einen Server gesendet. Verarbeitung erfolgt **lokal im Browser**. Optional können Sie pro Modus (**Kursteams**, **Jahrgangsgruppen**, **ARGEs**) einen getrennten Zwischenstand im **lokalen Speicher** dieses Browsers sichern (Schaltflächen oben in der App).

## Nutzung

1. Repository klonen oder die Dateien herunterladen.
2. `ms365-schooltool.html` im Browser öffnen **oder** ein statisches Hosting nutzen (z. B. GitHub Pages).
3. Oben zwischen **Kursteams**, **Jahrgangsgruppen** und **ARGEs** wählen.

### Modus A: Kursteams

1. Fächerliste exportieren und durch die Schritte in der App führen.
2. Entweder `neueteams.csv` + kurzes Skript aus der Anleitung **oder** die Datei **Kursteam-Anlage.cmd** herunterladen (enthält das PowerShell-Skript eingebettet) und per Doppelklick starten (Anmeldung interaktiv für MFA oder optional per `Get-Credential`).

### Modus B: Jahrgangsgruppen (Microsoft 365-Gruppen)

1. Domain und Präfix (z. B. `jg`) einstellen; Klassenzeilen im Format `1AK;2030` (Klasse;Abschlussjahr) eintragen.
2. Besitzer-UPNs pro Klasse eintragen.
3. Das generierte **Microsoft Graph PowerShell**-Skript kopieren **oder** **Jahrgangsgruppen-Anlage.cmd** herunterladen und starten (interaktive Graph-Anmeldung, MFA möglich).

Dokumentation: [New-MgGroup](https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup), [Gruppe erstellen (Graph)](https://learn.microsoft.com/graph/api/group-post-groups).

### Modus C: ARGEs (Microsoft 365-Gruppen)

1. **Domain** und optional **Präfix** für das Mail-Nickname festlegen (Präfix + aus dem Fachnamen erzeugter Teil, z. B. Fach `Deutsch` → `arge-deutsch`). Optional: Mail-Nickname in Großbuchstaben.
2. **Fächer / Bezeichnungen** eintragen: **eine Zeile pro ARGE** – z. B. aus Excel kopieren (`Deutsch`, `Mathematik`, `ARGE BB`, …). Daraus erzeugt die App den **Anzeigenamen** der Gruppe (`ARGE …`, sofern die Zeile nicht schon mit `ARGE ` beginnt) und das **Mail-Nickname** (Umlaute und Sonderzeichen werden für den technischen Teil normalisiert). Unter dem Eingabefeld zeigt eine **Live-Vorschau** als Tabelle: Anzeigename, Fach technisch, Mail-Nickname und die **E-Mail-Adresse** (`MailNickname@Domain`).
3. **Optional (fortgeschritten):** Weiterhin pro Zeile `Anzeigename;MailNickname` möglich – dann gilt der rechte Teil als festes Mail-Nickname (z. B. `ARGE BB;ARGEBB`).
4. Besitzer-UPNs pro ARGE eintragen.
5. Script kopieren oder **ARGE-Gruppen-Anlage.cmd** herunterladen und starten (interaktive Graph-Anmeldung, MFA möglich).

### Erwartete Spalten (flexibel)

Die App erkennt u. a.: **Klasse(n)**, **Fach**, **Lehrer**, **Schülergruppe** (je nach Export unterschiedlich benannt).

### Dateien

| Datei | Beschreibung |
|--------|----------------|
| `ms365-schooltool.html` | Hauptseite |
| `app.js` | Logik Kursteams (CSV-Export, Filter, lokaler Speicher, …) |
| `jahrgang.js` | Assistent Jahrgangsgruppen (Microsoft Graph PowerShell) |
| `arge.js` | Assistent ARGEs (Microsoft Graph PowerShell) |
| `polyglot-cmd.js` | Erzeugt eine **einzige** `.cmd`-Datei mit eingebettetem PowerShell (kein separates `.ps1` im Download) |
| `index.html` | Optional: Weiterleitung zur Hauptseite (für GitHub Pages-Start-URL) |

### Windows: Downloads und Sicherheit

Aus dem Browser heruntergeladene Dateien können mit **Mark of the Web** (Zone) markiert sein. Windows kann dann melden, dass die Datei „wegen Internetsicherheitseinstellungen“ nicht geöffnet werden darf – das betrifft auch **eine einzelne** `.cmd` im Ordner **Downloads**.

**So gehen Sie vor (privat / Einzelplatz):**

1. **Rechtsklick** auf die heruntergeladene `.cmd` → **Eigenschaften**.
2. Unten **Zulassen** (engl. **Unblock**) aktivieren → **OK**.
3. Datei **erneut** per Doppelklick starten.

**Alternative:** In **PowerShell** (als Benutzer reicht meist):

```powershell
Unblock-File -LiteralPath "$env:USERPROFILE\Downloads\ARGE-Gruppen-Anlage.cmd"
```

(Pfad und Dateiname anpassen.)

Bei **SmartScreen** („Windows hat den PC geschützt“): **Weitere Informationen** → **Trotzdem ausführen** – nur wenn Sie die Datei aus dieser App selbst erzeugt haben.

Die App liefert **eine** `.cmd` mit eingebettetem PowerShell (keine separate `.ps1` im Download-Ordner). Eine **100 % warnungsfreie** Ausführung aus dem Internet ohne Signatur oder IT-Freigabe kann Windows nicht garantieren.

## GitHub Pages

Repository auf **Pages** schalten (Branch `main`, Ordner `/`). Mit dem Repo-Namen **`ms365-schultools`** ist die App typischerweise unter:

- **Startseite:** [https://kurtsoeser.github.io/ms365-schultools/](https://kurtsoeser.github.io/ms365-schultools/)
- **Haupt-App:** [https://kurtsoeser.github.io/ms365-schultools/ms365-schooltool.html](https://kurtsoeser.github.io/ms365-schultools/ms365-schooltool.html)

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
