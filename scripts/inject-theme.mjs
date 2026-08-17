import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { join, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

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

const fouc =
    "<script>(function(){try{var k='ms365-theme-v1';var t=localStorage.getItem(k);if(t!=='dark'&&t!=='light'){t=(window.matchMedia&&matchMedia('(prefers-color-scheme: dark)').matches)?'dark':'light';}var d=document.documentElement;d.setAttribute('data-theme',t);d.style.colorScheme=t;}catch(e){}})();</script>";

const files = walkHtml(root);
let updated = 0;

for (const full of files) {
    const rel = relative(root, full).replace(/\\/g, '/');
    let html = readFileSync(full, 'utf8');
    let changed = false;

    if (!html.includes('ms365-theme-v1')) {
        if (/<head(\s[^>]*)?>/i.test(html)) {
            html = html.replace(/<head(\s[^>]*)?>/i, (m) => m + '\n    ' + fouc);
            changed = true;
        }
    }

    const depth = (rel.match(/\//g) || []).length;
    const src = (depth ? '../'.repeat(depth) : '') + 'src/shared/theme-toggle.js';
    if (!html.includes('theme-toggle.js')) {
        if (/msal-auth-ui\.js/.test(html)) {
            html = html.replace(
                /(<script[^>]*msal-auth-ui\.js[^>]*><\/script>)/i,
                '$1\n    <script src="' + src + '" defer></script>'
            );
            changed = true;
        } else if (html.includes('</body>')) {
            html = html.replace('</body>', '    <script src="' + src + '" defer></script>\n</body>');
            changed = true;
        }
    }

    if (changed) {
        writeFileSync(full, html);
        updated += 1;
        console.log('updated', rel);
    }
}

console.log('done, updated', updated, 'of', files.length);
