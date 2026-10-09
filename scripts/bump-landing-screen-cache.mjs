/**
 * Setzt ?v=YYYYMMDD auf allen Landing-Screenshot-URLs in landing/index.html.
 * Usage: node scripts/bump-landing-screen-cache.mjs
 */
import fs from 'node:fs';

const p = 'landing/index.html';
const stamp =
  process.env.MS365_SCREEN_CACHE_V ||
  new Date().toISOString().slice(0, 10).replace(/-/g, '');

let s = fs.readFileSync(p, 'utf8');
const re = /assets\/screens\/([a-z0-9-]+)\.png(\?v=[^"'&\s]*)?/g;
s = s.replace(re, `assets/screens/$1.png?v=${stamp}`);
fs.writeFileSync(p, s);
const n = (s.match(new RegExp(`screens/[a-z0-9-]+\\.png\\?v=${stamp}`, 'g')) || []).length;
console.log('bumped', n, 'screen urls to v=' + stamp);
