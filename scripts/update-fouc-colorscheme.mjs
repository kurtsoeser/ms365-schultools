import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { join, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const needle = "document.documentElement.setAttribute('data-theme',t);";
const insert = "var d=document.documentElement;d.setAttribute('data-theme',t);d.style.colorScheme=t;";

function walkHtml(dir, out = []) {
    for (const name of readdirSync(dir)) {
        if (name === 'node_modules' || name === 'dist' || name === '.git') continue;
        const full = join(dir, name);
        const st = statSync(full);
        if (st.isDirectory()) walkHtml(full, out);
        else if (name.endsWith('.html')) out.push(full);
    }
    return out;
}

let n = 0;
for (const full of walkHtml(root)) {
    let html = readFileSync(full, 'utf8');
    if (!html.includes(needle) || html.includes('d.style.colorScheme')) continue;
    writeFileSync(full, html.replace(needle, insert));
    n += 1;
    console.log(relative(root, full).replace(/\\/g, '/'));
}
console.log('updated', n);
