import { defineConfig } from 'vite';
import { readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { appBuildInfoPlugin } from './scripts/app-build-info.mjs';

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

export default defineConfig(() => {
  // For GitHub Pages Project Pages set VITE_BASE="/<repo-name>/"
  const base = withTrailingSlash(process.env.VITE_BASE || '/');

  return {
    base,
    plugins: [appBuildInfoPlugin()],
    build: {
      outDir: 'dist',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          welcome: resolve(__dirname, 'welcome.html'),
          index: resolve(__dirname, 'index.html'),
          schooltool: resolve(__dirname, 'ms365-schooltool.html'),
          tenant: resolve(__dirname, 'tenant.html'),
          einrichtung: resolve(__dirname, 'einrichtung.html'),
          ersteinrichtung: resolve(__dirname, 'ersteinrichtung.html'),
          help: resolve(__dirname, 'hilfe.html'),
          ...htmlEntriesFrom('tools'),
          ...htmlEntriesFrom('tools/archiv')
        }
      }
    }
  };
});

