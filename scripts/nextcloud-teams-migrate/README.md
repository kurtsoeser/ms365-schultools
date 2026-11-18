# Nextcloud → Teams Datei-Migration

Lokales CLI: Dateien per **WebDAV** aus Nextcloud lesen und per **Microsoft Graph** in Team-Dateien hochladen. Für große Mengen (mehrere GB) gedacht – nicht im Browser.

**Grafisches Werkzeug** in der App: [`tools/nextcloud-teams-migrate.html`](../../tools/nextcloud-teams-migrate.html) (Ordnerbrowser, Dry-Run, Kanal-Mapping). Das CLI bleibt die robuste Variante für 10+ GB.

Pilot: Fachschaft **Physik**.

## Grafisch in der App (empfohlen)

1. `.env.local` aus `.env.example` anlegen (Benutzer + App-Passwort)
2. **`Start-Helfer.cmd` doppelklicken** (Fenster offen lassen)
3. In der App **Nextcloud → Teams** → Reiter **Lokaler Helfer** → **Verbinden**
4. Ordner browsen, Team wählen, Dry-Run, Übertragen

Der Helfer läuft nur auf `http://127.0.0.1:3791` und umgeht CORS.

## Voraussetzungen

1. Nextcloud **App-Passwort** (Personal → Security), z. B. Name `Teams-Migration`
2. Konto mit Zugriff auf den Quellordner in Nextcloud
3. Microsoft-365-Konto, das Dateien ins Ziel-Team schreiben darf
4. Für CLI-Upload: Entra-App mit Device-Code; in der App: normale MSAL-Anmeldung

## WebDAV-Adresse

```text
https://www.dsrom-intranet.de/remote.php/dav/files/<NEXTCLOUD_USER>/
```

Beispiel Physik (Pfad anpassen, falls die Hierarchie abweicht):

```text
…/files/<USER>/5. Gymnasium (5-12)/1. Fächer/Physik/
```

Login: `NEXTCLOUD_USER` + App-Passwort.

## Setup

```bash
cd scripts/nextcloud-teams-migrate
cp .env.example .env.local
# .env.local ausfüllen
```

Mapping prüfen/anpassen: `mapping.physik.json`  
(`teamName` oder feste `groupId` des Teams)

## Befehle

Alle Befehle aus dem Repo-Root oder diesem Ordner:

```bash
# Verbindungstest
node scripts/nextcloud-teams-migrate/migrate.mjs check

# Inventar (Anzahl + Größe)
node scripts/nextcloud-teams-migrate/migrate.mjs inventory --path "5. Gymnasium (5-12)/1. Fächer/Physik"

# Dry-Run: Mapping + Dateiliste, kein Upload
node scripts/nextcloud-teams-migrate/migrate.mjs dry-run --mapping mapping.physik.json

# Smoke-Test: max. 5 Dateien
node scripts/nextcloud-teams-migrate/migrate.mjs migrate --mapping mapping.physik.json --limit 5

# Smoke-Test: max. ~50 MB
node scripts/nextcloud-teams-migrate/migrate.mjs migrate --mapping mapping.physik.json --max-bytes 52428800

# Voller Physik-Lauf (Resume-fähig)
node scripts/nextcloud-teams-migrate/migrate.mjs migrate --mapping mapping.physik.json
```

Bei `migrate` erscheint ein Device-Code für die Microsoft-Anmeldung.

## Zielablage

Im Team-Dateien-Stamm:

```text
Import aus Nextcloud/Physik/…
```

Struktur unter dem Nextcloud-Physik-Ordner wird 1:1 übernommen.  
`DEST_ROOT_FOLDER` in `.env.local` änderbar.

## Resume

Erfolgreiche Dateien landen unter `.state/done-*.json`.  
Abbruch → denselben Befehl erneut starten; bereits übertragene Dateien werden übersprungen.

## Weitere Fachschaften

Weitere Einträge in einer Mapping-JSON analog zu `mapping.physik.example.json` ergänzen und nacheinander migrieren (Last auf Nextcloud schonen).

## Nicht enthalten

- Live-Sync / Delta
- Nextcloud-Freigaberechte
- Chat / Planner
- UI in MS365-Schul-Tools
