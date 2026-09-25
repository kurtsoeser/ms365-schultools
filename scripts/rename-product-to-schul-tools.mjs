/**
 * Rename product brand: MS365-Schul-Tools → MS365-Schul-Tools
 * Does NOT touch generic "Schulverwaltung" (school admin role/department).
 */
import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const skipDirs = new Set(['node_modules', 'dist', '.git', '.cursor']);
const exts = new Set([
  '.html',
  '.js',
  '.mjs',
  '.css',
  '.md',
  '.json',
  '.yml',
  '.yaml',
  '.txt',
  '.svg'
]);

const replacements = [
  [/MS365-Schul-Tools/g, 'MS365-Schul-Tools'],
  [/MS365-Schul-Tools/g, 'MS365-Schul-Tools']
];

function walk(dir, out = []) {
  for (const name of fs.readdirSync(dir)) {
    if (skipDirs.has(name)) continue;
    const full = path.join(dir, name);
    const st = fs.statSync(full);
    if (st.isDirectory()) walk(full, out);
    else if (exts.has(path.extname(name).toLowerCase())) out.push(full);
  }
  return out;
}

let filesChanged = 0;
let hits = 0;
for (const file of walk(root)) {
  let s = fs.readFileSync(file, 'utf8');
  let n = 0;
  for (const [re, to] of replacements) {
    const before = s;
    s = s.replace(re, to);
    if (s !== before) {
      const m = before.match(re);
      n += m ? m.length : 0;
    }
  }
  if (n > 0) {
    fs.writeFileSync(file, s);
    filesChanged += 1;
    hits += n;
    console.log(path.relative(root, file).replace(/\\/g, '/'), n);
  }
}
console.log('done:', filesChanged, 'files,', hits, 'replacements');
