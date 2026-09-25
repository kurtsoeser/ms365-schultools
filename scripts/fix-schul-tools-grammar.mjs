import fs from 'node:fs';
import path from 'node:path';

const fixes = [
  ['Willkommen bei MS365-Schul-Tools', 'Willkommen bei MS365-Schul-Tools'],
  ['Dashboard von MS365-Schul-Tools', 'Dashboard von MS365-Schul-Tools'],
  ['in MS365-Schul-Tools', 'in MS365-Schul-Tools'],
  ['Anbindung von MS365-Schul-Tools', 'Anbindung von MS365-Schul-Tools'],
  ['Backup von MS365-Schul-Tools', 'Backup von MS365-Schul-Tools'],
  ['Landing Page – MS365-Schul-Tools', 'Landing Page – MS365-Schul-Tools']
];

const skip = new Set(['node_modules', 'dist', '.git']);

function walk(dir, out = []) {
  for (const name of fs.readdirSync(dir)) {
    if (skip.has(name)) continue;
    const full = path.join(dir, name);
    const st = fs.statSync(full);
    if (st.isDirectory()) walk(full, out);
    else if (/\.(html|js|mjs|md)$/i.test(name)) out.push(full);
  }
  return out;
}

let changed = 0;
for (const file of walk(process.cwd())) {
  let s = fs.readFileSync(file, 'utf8');
  let hit = false;
  for (const [from, to] of fixes) {
    if (s.includes(from)) {
      s = s.split(from).join(to);
      hit = true;
    }
  }
  if (hit) {
    fs.writeFileSync(file, s);
    changed += 1;
    console.log(path.relative(process.cwd(), file).replace(/\\/g, '/'));
  }
}
console.log('fixed', changed, 'files');
