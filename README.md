# WebUntis → Microsoft Teams (Kursteam-Creator)

Einfache **reine Browser-App** (ohne Server): WebUntis-Exporte (CSV/Excel) werden aufbereitet, damit Sie daraus **Microsoft-Kursteams** per **PowerShell** (`New-Team -Template "EDU_Class"`) anlegen können.

## Datenschutz

Es werden **keine Daten** an einen Server gesendet. Verarbeitung erfolgt **lokal im Browser**. Optional können Sie einen Zwischenstand im **lokalen Speicher** dieses Browsers sichern.

## Nutzung

1. Repository klonen oder die Dateien herunterladen.
2. `webuntis-teams-creator.html` im Browser öffnen **oder** ein statisches Hosting nutzen (z. B. GitHub Pages).
3. WebUntis-Fächerliste exportieren und durch die Schritte in der App führen.
4. `neueteams.csv` herunterladen und das angezeigte PowerShell-Skript im gleichen Ordner ausführen (Microsoft Teams PowerShell, angemeldeter Admin).

### Erwartete Spalten (flexibel)

Die App erkennt u. a.: **Klasse(n)**, **Fach**, **Lehrer**, **Schülergruppe** (je nach Export unterschiedlich benannt).

### Dateien

| Datei | Beschreibung |
|--------|----------------|
| `webuntis-teams-creator.html` | Hauptseite |
| `app.js` | Logik (CSV-Export, Filter, lokaler Speicher, …) |
| `index.html` | Optional: Weiterleitung zur Hauptseite (für GitHub Pages-Start-URL) |

## GitHub Pages

Repository auf **Pages** schalten (Branch `main`, Ordner `/`). Die App ist dann unter  
`https://<user>.github.io/<repo>/` erreichbar – mit `index.html` als Einstieg oder direkt  
`…/webuntis-teams-creator.html`.

**Hinweis:** Bei kostenlosem GitHub ist Pages für **private** Repos oft nicht verfügbar; öffentliches Repo oder GitHub Pro nötig.

## Lizenz

Keine Lizenz gesetzt – ergänzen Sie bei Bedarf eine `LICENSE` im Repository.
