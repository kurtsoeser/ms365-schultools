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

/**
 * Theme-Boot als externes Skript (kein Inline-Skript mehr – die Content-Security-Policy
 * erlaubt nur script-src 'self'). admin.html hat eine eigene Variante.
 */
function foucScript(isAdmin, prefix) {
    const theme = '<script src="' + prefix + 'src/shared/theme-boot.js?v=4"></script>';
    if (!isAdmin) return theme;
    return theme + '\n    <script src="' + prefix + 'src/pages/admin-boot.js"></script>';
}

const oldFoucRe =
    /<script>\(function\(\)\{try\{[^<]*ms365-theme-v1[^<]*\}\)\(\);<\/script>/;
const bootScriptRe = /src\/(?:shared\/theme-boot|pages\/(?:index|admin)-boot)\.js/;

const files = walkHtml(root);
let updated = 0;

for (const full of files) {
    const rel = relative(root, full).replace(/\\/g, '/');
    if (rel.startsWith('landing/') || rel.startsWith('dist/')) continue;
    let html = readFileSync(full, 'utf8');
    let changed = false;
    const isAdmin = rel === 'admin.html' || rel.endsWith('/admin.html');
    const depth = (rel.match(/\//g) || []).length;
    const prefix = depth ? '../'.repeat(depth) : '';

    if (oldFoucRe.test(html)) {
        html = html.replace(oldFoucRe, foucScript(isAdmin, prefix));
        changed = true;
    } else if (!bootScriptRe.test(html)) {
        if (/<head(\s[^>]*)?>/i.test(html)) {
            html = html.replace(/<head(\s[^>]*)?>/i, (m) => m + '\n    ' + foucScript(isAdmin, prefix));
            changed = true;
        }
    }

    const src = prefix + 'src/shared/theme-toggle.js?v=4';
    if (/theme-toggle\.js\?v=[123]/.test(html)) {
        html = html.replace(/theme-toggle\.js\?v=[123]/g, 'theme-toggle.js?v=4');
        changed = true;
    } else if (html.includes('src/shared/theme-toggle.js"') && !html.includes('theme-toggle.js?v=')) {
        html = html.replace(/src\/shared\/theme-toggle\.js/g, 'src/shared/theme-toggle.js?v=4');
        changed = true;
    }
    if (html.includes('app.css?v=12')) {
        html = html.replace(/app\.css\?v=12/g, 'app.css?v=13');
        changed = true;
    }
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
