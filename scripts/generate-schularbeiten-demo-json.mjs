/**
 * Demo-Daten Schularbeiten 2026/27 als JSON (für Dokumentation / Stammdaten-Import).
 * Generiert aus schularbeiten-planer-demo-data.js – bei Änderungen neu erzeugen:
 *   node scripts/generate-schularbeiten-demo-json.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, '..');
const modPath = resolve(root, 'src/tools/schularbeiten-planer/schularbeiten-planer-demo-data.js');
const { getDemoSeedPackage } = await import(pathToFileURL(modPath).href);
const pack = getDemoSeedPackage();

const outDir = resolve(root, 'docs/demo-data');
mkdirSync(outDir, { recursive: true });
const outFile = resolve(outDir, 'schularbeiten-2026-27.json');
writeFileSync(outFile, JSON.stringify(pack, null, 2), 'utf8');
console.log('Wrote', outFile);
console.log('Counts:', pack.counts);
