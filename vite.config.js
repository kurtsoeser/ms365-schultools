import { defineConfig } from 'vite';
import { createReadStream, existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { extname, join, normalize, resolve, sep } from 'node:path';
import { appBuildInfoPlugin } from './scripts/app-build-info.mjs';
import { cspMetaPlugin } from './scripts/csp-policy.mjs';

function withTrailingSlash(value) {
  if (!value) return '/';
  return value.endsWith('/') ? value : `${value}/`;
}

/** Alle HTML-Seiten eines Ordners als Vite-MPA-Entries (sonst 404 auf GitHub Pages). */
function htmlEntriesFrom(relDir) {
  const absDir = resolve(__dirname, relDir);
  const entries = {};
  let names;
  try {
    names = readdirSync(absDir);
  } catch {
    return entries;
  }
  for (const name of names) {
    if (!name.endsWith('.html')) continue;
    const rel = `${relDir}/${name}`.replace(/\\/g, '/');
    const key = rel.replace(/[^a-zA-Z0-9]+/g, '_');
    entries[key] = resolve(absDir, name);
  }
  return entries;
}

const LANDING_MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff2': 'font/woff2'
};

/** Dev-Server: /landing → Ordner landing/ (auch im Build per copy-static). */
function serveLandingPlugin() {
  const landingRoot = resolve(__dirname, 'landing');
  return {
    name: 'serve-landing',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const raw = (req.url || '').split('?')[0];
        if (!raw.startsWith('/landing')) return next();
        let rel = decodeURIComponent(raw.slice('/landing'.length) || '/');
        if (rel === '/' || rel === '') rel = '/index.html';
        const filePath = normalize(join(landingRoot, rel.replace(/^\//, '')));
        if (!filePath.startsWith(landingRoot + sep) && filePath !== landingRoot) {
          res.statusCode = 403;
          res.end('Forbidden');
          return;
        }
        if (!existsSync(filePath) || !statSync(filePath).isFile()) return next();
        res.setHeader(
          'Content-Type',
          LANDING_MIME[extname(filePath).toLowerCase()] || 'application/octet-stream'
        );
        createReadStream(filePath).pipe(res);
      });
    }
  };
}

/** Nur Dev: Admin-Key aus gitignored Datei für Pilot-Verdrahtung. */
function serveSspDevAdminKeyPlugin() {
  const keyPath = resolve(__dirname, 'backend/stammdaten-portal-api/.admin-key.local.txt');
  return {
    name: 'ssp-dev-admin-key',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const raw = (req.url || '').split('?')[0];
        if (raw !== '/__ssp-dev-admin-key') return next();
        if (!existsSync(keyPath)) {
          res.statusCode = 404;
          res.setHeader('Content-Type', 'application/json; charset=utf-8');
          res.end(JSON.stringify({ error: 'no local admin key file' }));
          return;
        }
        const key = readFileSync(keyPath, 'utf8').trim();
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.setHeader('Cache-Control', 'no-store');
        res.end(JSON.stringify({ adminKey: key }));
      });
    }
  };
}

export default defineConfig(() => {
  // For GitHub Pages Project Pages set VITE_BASE="/<repo-name>/"
  const base = withTrailingSlash(process.env.VITE_BASE || '/');

  return {
    base,
    plugins: [appBuildInfoPlugin(), serveLandingPlugin(), serveSspDevAdminKeyPlugin(), cspMetaPlugin()],
    server: {
      // Stammdaten-Portal-Gateway: lokaler Proxy → kein Browser-CORS gegen Azure
      proxy: {
        '/__ssp-gw': {
          target: 'https://func-ms365-ssp-dev.azurewebsites.net',
          changeOrigin: true,
          secure: true,
          rewrite: (path) => path.replace(/^\/__ssp-gw/, '')
        }
      }
    },
    preview: {
      proxy: {
        '/__ssp-gw': {
          target: 'https://func-ms365-ssp-dev.azurewebsites.net',
          changeOrigin: true,
          secure: true,
          rewrite: (path) => path.replace(/^\/__ssp-gw/, '')
        }
      }
    },
    build: {
      outDir: 'dist',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          // Root-HTMLs automatisch (sonst 404 auf GitHub Pages, z. B. action-log.html)
          ...htmlEntriesFrom('.'),
          ...htmlEntriesFrom('tools'),
          ...htmlEntriesFrom('tools/archiv')
        }
      }
    }
  };
});
