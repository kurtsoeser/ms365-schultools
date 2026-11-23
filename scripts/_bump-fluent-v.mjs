import fs from 'fs';
const p = new URL('../src/shared/app-header-chrome.js', import.meta.url);
let s = fs.readFileSync(p, 'utf8');
const m = s.match(/fluent-blocks\.css[^,\n]*/);
console.log('match:', m && m[0]);
s = s.replace(/fluent-blocks\.css'\) \+ '\?v=\d+'/g, "fluent-blocks.css') + '?v=6'");
fs.writeFileSync(p, s);
console.log('done');
