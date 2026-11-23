/**
 * Schreibt landing/screens-catalog.js aus scripts/landing-screens-data.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { SCREEN_PERSONAS } from './landing-screens-data.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outFile = path.join(__dirname, '..', 'landing', 'screens-catalog.js');

const publicPersonas = SCREEN_PERSONAS.map((p) => ({
  id: p.id,
  label: p.label,
  short: p.short,
  lead: p.lead,
  screens: p.screens.map((s) => ({
    id: s.id,
    file: s.file,
    label: s.label,
    urlLabel: s.urlLabel,
    caption: s.caption
  }))
}));

const body =
  '/**\n' +
  ' * Auto-generated – Quelle: scripts/landing-screens-data.mjs\n' +
  ' * Neu: node scripts/sync-landing-screens-catalog.mjs\n' +
  ' */\n' +
  'window.MS365_SCREENS_CATALOG = ' +
  JSON.stringify({ personas: publicPersonas }, null, 2) +
  ';\n';

fs.writeFileSync(outFile, body, 'utf8');
const n = publicPersonas.reduce((a, p) => a + p.screens.length, 0);
console.log('wrote landing/screens-catalog.js ·', publicPersonas.length, 'personas ·', n, 'screens');
