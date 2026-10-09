# Veröffentlichungs-Durchlauf

Ein Push auf `main` startet automatisch CI und Deploy. **Vorher** lokal vorbereiten:

```bash
npm run release:prep
```

Das Skript führt aus:

1. Secret-Check (`check:secrets`)
2. Landing-Screenshots (startet kurz Vite auf Port 5199, Playwright muss installiert sein: `npx playwright install chromium`)
3. Screenshot-Cache in `landing/index.html` (`?v=YYYYMMDD`)
4. Footer **„Homepage zuletzt aktualisiert“** + `landing/site-build.json` (Skript `write-landing-build-info.mjs`)
5. Release-Notes aus letzten Commits (`public/release-notes.json`)
6. `npm test`, `npm run lint`, `npm run build`

Optionen:

```bash
npm run release:prep -- --no-screens
npm run release:prep -- --no-notes
```

Danach committen und pushen:

```bash
git add -A
git commit -m "Release: …"
git push origin main
```

## Cursor / Agent

Im Chat reicht z. B.:

- **„Veröffentlichungs-Durchlauf“**
- **„/release“**
- **„alles veröffentlichen wie beim letzten Mal“**

Die Regel `.cursor/rules/release-veroeffentlichung.mdc` beschreibt den vollen Ablauf für den Agenten.

## Was nach dem Push passiert

| Workflow | Ziel |
|----------|------|
| `ci.yml` | Test, Lint, Build |
| `deploy-ftp.yml` | `app.ms365.schule` |
| `deploy-pages.yml` | GitHub Pages (Build-Artefakt) |
| `deploy-landing-ftp.yml` | **`https://ms365.schule/`** (nur Ordner `landing/`) |

### Wichtig: zwei FTP-Ziele

| Domain | Secret | Inhalt |
|--------|--------|--------|
| `app.ms365.schule` | `FTP_REMOTE_DIR` (z. B. `/app/`) | `dist/` nach Build |
| **`ms365.schule`** | **`FTP_LANDING_REMOTE_DIR`** (z. B. `/` oder Webroot) | Repo-Ordner **`landing/`** |

Ohne `FTP_LANDING_REMOTE_DIR` in den GitHub-Repo-Secrets bleibt die Marketing-Homepage alt (nur `app.ms365.schule` wird aktualisiert).

**Prüfen, ob live:** Footer „Homepage zuletzt aktualisiert“ oder Seitenquelltext `meta name="ms365-landing-build"`.

Lokal hochladen (mit `.ftp-credentials.local` inkl. `FTP_LANDING_REMOTE_DIR`):

```powershell
powershell -File scripts/ftp-upload-landing.ps1
```
