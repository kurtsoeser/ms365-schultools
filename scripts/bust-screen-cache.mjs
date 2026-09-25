import fs from 'node:fs';

const p = 'landing/index.html';
let s = fs.readFileSync(p, 'utf8');

// Normalize any doubled or old bust params to a single v=name1
s = s.replace(/assets\/screens\/([a-z0-9-]+)\.png(\?v=[^"'&\s]*)?/g, 'assets/screens/$1.png?v=name1');

fs.writeFileSync(p, s);
const n = (s.match(/screens\/[a-z0-9-]+\.png\?v=name1/g) || []).length;
console.log('normalized', n, 'urls');
