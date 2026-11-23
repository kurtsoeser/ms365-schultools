import fs from 'fs';
const p = new URL('../src/shared/app-header-chrome.js', import.meta.url);
let s = fs.readFileSync(p, 'utf8');
s = s.replace(/shell-fluent\.css'\) \+ '\?v=\d+'/g, "shell-fluent.css') + '?v=9'");
fs.writeFileSync(p, s);
console.log('ok', !!s.includes("shell-fluent.css') + '?v=9'"));
