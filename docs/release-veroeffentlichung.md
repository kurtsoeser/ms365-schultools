# Veröffentlichungs-Durchlauf

Ein Push auf `main` startet automatisch CI und Deploy. **Vorher** lokal vorbereiten:

```bash
npm run release:prep
```

Das Skript führt aus:

1. Secret-Check (`check:secrets`)
2. Landing-Screenshots (startet kurz Vite auf Port 5199, Playwright muss installiert sein: `npx playwright install chromium`)
3. Screenshot-Cache in `landing/index.html` (`?v=YYYYMMDD`)
4. Release-Notes aus letzten Commits (`public/release-notes.json`)
5. `npm test`, `npm run lint`, `npm run build`

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

Marketing-Homepage unter `ms365.schule` nutzt die Dateien aus `landing/` (per FTP/Hosting des Anbieters – ggf. `landing/` separat hochladen, wenn nicht über `dist/landing`).
