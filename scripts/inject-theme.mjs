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

function foucScript(withAdminBootClear) {
    const clear = withAdminBootClear ? "d.removeAttribute('data-ms365-admin-boot');" : '';
    return (
        "<script>(function(){try{var d=document.documentElement;" +
        "var t=localStorage.getItem('ms365-theme-v1');if(t!=='dark'&&t!=='light'){t='light';}" +
        "d.setAttribute('data-theme',t);d.style.colorScheme=t;" +
        "var b=localStorage.getItem('ms365-brand-v1');if(b!=='classic'&&b!=='teal'){b='teal';}" +
        "d.setAttribute('data-brand',b);" +
        clear +
        "}catch(e){}})();</script>"
    );
}

const oldFoucRe =
    /<script>\(function\(\)\{try\{[^<]*ms365-theme-v1[^<]*\}\)\(\);<\/script>/;

const files = walkHtml(root);
let updated = 0;

for (const full of files) {
    const rel = relative(root, full).replace(/\\/g, '/');
    if (rel.startsWith('landing/') || rel.startsWith('dist/')) continue;
    let html = readFileSync(full, 'utf8');
    let changed = false;
    const isAdmin = rel === 'admin.html' || rel.endsWith('/admin.html');

    if (oldFoucRe.test(html)) {
        html = html.replace(oldFoucRe, foucScript(isAdmin));
        changed = true;
    } else if (!html.includes('ms365-theme-v1')) {
        if (/<head(\s[^>]*)?>/i.test(html)) {
            html = html.replace(/<head(\s[^>]*)?>/i, (m) => m + '\n    ' + foucScript(isAdmin));
            changed = true;
        }
    } else if (!html.includes('ms365-brand-v1')) {
        if (html.includes("d.setAttribute('data-theme',t);d.style.colorScheme=t;")) {
            html = html.replace(
                "d.setAttribute('data-theme',t);d.style.colorScheme=t;",
                "d.setAttribute('data-theme',t);d.style.colorScheme=t;var b=localStorage.getItem('ms365-brand-v1');if(b!=='classic'&&b!=='teal'){b='teal';}d.setAttribute('data-brand',b);"
            );
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
