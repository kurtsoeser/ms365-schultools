import fs from 'node:fs/promises';
import path from 'node:path';

const projectRoot = path.resolve(process.cwd());
const distRoot = path.resolve(projectRoot, 'dist');

async function exists(p) {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

async function copyFileToDist(relFile) {
  const src = path.resolve(projectRoot, relFile);
  const dst = path.resolve(distRoot, relFile);
  if (!(await exists(src))) return;
  await fs.mkdir(path.dirname(dst), { recursive: true });
  await fs.copyFile(src, dst);
}

async function copyDirToDist(relDir) {
  const src = path.resolve(projectRoot, relDir);
  const dst = path.resolve(distRoot, relDir);
  if (!(await exists(src))) return;
  await fs.mkdir(path.dirname(dst), { recursive: true });
  await fs.cp(src, dst, { recursive: true });
}

async function main() {
  if (!(await exists(distRoot))) {
    throw new Error('dist/ fehlt – zuerst `vite build` ausführen.');
  }

  // Ensure ms365-schooltool.html is always a small redirect page.
  // (The actual tools live in index.html + tools/*.html)
  const schooltoolRedirectHtml = `<!DOCTYPE html>
<html lang="de">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="ms365-graph-client-id" content="e1d877c3-004c-4040-8c3b-81a59e0c7050" />
    <title>MS365-Schulverwaltung</title>
    <link rel="stylesheet" href="app.css" />
    <script>
      (function () {
        const params = new URLSearchParams(window.location.search);
        const mode = String(params.get('mode') || '').toLowerCase().trim();
        const map = {
          kursteams: 'tools/kursteams.html',
          kursteam: 'tools/kursteams.html',
          webuntis: 'tools/kursteams.html',
          jahrgang: 'tools/jahrgang.html',
          arge: 'tools/arge.html',
          gruppenerstellung: 'tools/gruppenerstellung.html',
          grouppolicy: 'tools/gruppenerstellung.html'
        };
        window.location.replace(map[mode] || 'index.html');
      })();
    </script>
    <noscript><meta http-equiv="refresh" content="0;url=index.html" /></noscript>
  </head>
  <body>
    <div class="container" style="max-width: 1100px">
      <div class="header">
        <h1>MS365-Schulverwaltung</h1>
        <p>Weiterleitung…</p>
        <p class="header-help-row"><a href="hilfe.html" class="header-help-link">Hilfe &amp; Datenschutz</a></p>
      </div>
      <div class="content" style="padding-bottom: 30px">
        <p style="color: #6c757d; line-height: 1.5">
          Falls die Weiterleitung nicht funktioniert, öffnen Sie:
          <a href="index.html">Dashboard</a>, <a href="tools/kursteams.html">Kursteams</a>,
          <a href="tools/jahrgang.html">Jahrgangsgruppen</a>, <a href="tools/arge.html">ARGEs</a>,
          <a href="tools/gruppenerstellung.html">Gruppenerstellung</a>.
        </p>
      </div>
    </div>
  </body>
</html>
`;
  await fs.writeFile(path.resolve(distRoot, 'ms365-schooltool.html'), schooltoolRedirectHtml, 'utf8');

  // Root-level static files referenced by HTML
  const rootFiles = [
    'app.css',
    'app.js',
    'ms365-config.js',
    'ms365-config.example.js',
    'school-domain.js',
    'tenant-settings.js',
    'tenant-settings-core.js',
    'tenant-settings-ui.js',
    'polyglot-cmd.js',
    'jahrgang.js',
    'arge.js',
    'arge-graph.js',
    'kursteam-graph.js',
    'gruppenerstellung-policy.js',
    'README.md'
  ];

  for (const f of rootFiles) await copyFileToDist(f);

  // Script folder(s)
  await copyDirToDist('kursteam');
  await copyDirToDist('src');
}

await main();

