/**
 * Baut ein Self-Host-ZIP für die Schulhomepage (Formular + Module + Config).
 * Erwartet globales JSZip (CDN).
 */

import { exportSettingsAsPublicConfig, loadPortalSettings } from './stammdaten-portal-store.js';

const MODULE_FILES = [
    'stammdaten-portal-logic.js',
    'stammdaten-portal-fields.js',
    'stammdaten-portal-export.js',
    'stammdaten-portal-config.js',
    'stammdaten-portal-store.js',
    'stammdaten-portal-api.js',
    'stammdaten-portal-form-ui.js',
    'stammdaten-portal.css',
    'einschreiber-vorgaben.json'
];

/**
 * @param {{ JSZip?: any, settings?: object, fetchImpl?: typeof fetch }} [opts]
 * @returns {Promise<Blob>}
 */
export async function buildSelfHostZip(opts = {}) {
    const JSZip = opts.JSZip || globalThis.JSZip;
    if (!JSZip) throw new Error('JSZip nicht geladen.');
    const fetchImpl = opts.fetchImpl || fetch;
    const settings = opts.settings || loadPortalSettings();
    const zip = new JSZip();
    const root = zip.folder('schueler-stammdaten-formular');
    const modDir = root.folder('src/tools/schueler-stammdaten-portal');

    const base = new URL('../../tools/schueler-stammdaten-portal/', import.meta.url);
    // In Browser: resolve against current page
    const pageBase = typeof window !== 'undefined' ? window.location.href : base.href;

    async function getText(relFromToolsPortal) {
        const url = new URL(
            `../src/tools/schueler-stammdaten-portal/${relFromToolsPortal}`,
            pageBase.includes('/tools/') ? pageBase : new URL('tools/schueler-stammdaten-portal.html', pageBase).href
        );
        // Prefer absolute from app root when on admin page
        const candidates = [
            new URL(`../src/tools/schueler-stammdaten-portal/${relFromToolsPortal}`, pageBase).href,
            new URL(`/src/tools/schueler-stammdaten-portal/${relFromToolsPortal}`, window.location.origin).href
        ];
        let lastErr;
        for (const href of candidates) {
            try {
                const res = await fetchImpl(href, { cache: 'no-store' });
                if (res.ok) return await res.text();
                lastErr = new Error(`${href} → ${res.status}`);
            } catch (e) {
                lastErr = e;
            }
        }
        throw lastErr || new Error(relFromToolsPortal);
    }

    for (const file of MODULE_FILES) {
        modDir.file(file, await getText(file));
    }

    // Minimal CSS shell
    let appCss = '';
    try {
        const res = await fetchImpl(new URL('../app.css', pageBase).href, { cache: 'no-store' });
        if (res.ok) appCss = await res.text();
    } catch {
        appCss = 'body{font-family:system-ui,sans-serif;margin:0;padding:1rem;background:#f7f7f5;color:#1a1a1a}';
    }
    root.file('app.css', appCss);

    const cfg = {
        ...exportSettingsAsPublicConfig(settings),
        mode: settings.mode === 'remote' ? 'remote' : 'local',
        apiBaseUrl: settings.apiBaseUrl || ''
    };
    root.file(
        'stammdaten-portal-config.js',
        `window.MS365_STAMMDATEN_PORTAL = ${JSON.stringify(cfg, null, 2)};\n`
    );

    root.file(
        'index.html',
        `<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="robots" content="noindex,nofollow" />
  <title>Schüler-Stammdaten – ${escapeHtml(cfg.schoolName || 'Schule')}</title>
  <link rel="stylesheet" href="./app.css" />
  <link rel="stylesheet" href="./src/tools/schueler-stammdaten-portal/stammdaten-portal.css" />
  <script src="./stammdaten-portal-config.js"></script>
</head>
<body>
  <div class="container" style="max-width:800px;margin:0 auto;padding:1rem">
    <header style="margin-bottom:1rem">
      <h1 style="margin:0 0 .35rem;font-size:1.25rem">Schüler-Stammdaten</h1>
      <p style="margin:0;opacity:.75">Geschützte Eingabe für Erziehungsberechtigte – kein Microsoft-Konto nötig.</p>
    </header>
    <div id="sspFormApp" aria-live="polite"></div>
    <p style="margin-top:2rem;font-size:.8rem;opacity:.55">Bereitgestellt mit MS365-Schul-Tools</p>
  </div>
  <script type="module">
    import { mountParentForm } from './src/tools/schueler-stammdaten-portal/stammdaten-portal-form-ui.js';
    mountParentForm(document.getElementById('sspFormApp'));
  </script>
</body>
</html>
`
    );

    root.file(
        'README.txt',
        `Schüler-Stammdaten-Formular – Self-Host
=====================================

1) Ordner auf die Schulhomepage / Static Web App kopieren.
2) stammdaten-portal-config.js prüfen:
   - schoolName, mode ("remote"), apiBaseUrl (Gateway …/api)
   - KEIN Function Key, KEIN PIN-Pepper hier!
3) Eltern-Link: https://ihre-schule.at/.../index.html?id=<GUID>
   PIN getrennt übermitteln.
4) CORS am Gateway muss die Schuldomain erlauben.
5) Admin bleibt in MS365-Schul-Tools (Datensätze / Export).

Erzeugt: ${new Date().toISOString()}
`
    );

    return zip.generateAsync({ type: 'blob' });
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}
