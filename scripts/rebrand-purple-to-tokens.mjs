/**
 * Replace hardcoded classic purple brand colors with CSS brand tokens.
 * Skips: landing/, node_modules, dist, .git
 * Skips in app.css: html[data-brand="classic"] blocks, brand-swatch--classic
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SKIP_DIRS = new Set(['node_modules', 'dist', '.git', 'landing']);
const EXTENSIONS = new Set(['.html', '.css', '.js']);

function walkDir(dir, files = []) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        if (SKIP_DIRS.has(entry.name)) continue;
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            walkDir(full, files);
        } else if (EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
            files.push(full);
        }
    }
    return files;
}

function rgbaToColorMix(alpha, brandVar) {
    const percent = Math.round(parseFloat(alpha) * 100);
    return `color-mix(in srgb, ${brandVar} ${percent}%, transparent)`;
}

function replaceRgba(content) {
    return content
        .replace(
            /rgba\s*\(\s*94\s*,\s*114\s*,\s*228\s*,\s*([\d.]+)\s*\)/gi,
            (_, a) => rgbaToColorMix(a, 'var(--brand1)')
        )
        .replace(
            /rgba\s*\(\s*130\s*,\s*94\s*,\s*228\s*,\s*([\d.]+)\s*\)/gi,
            (_, a) => rgbaToColorMix(a, 'var(--brand2)')
        );
}

function replaceHex(content) {
    const hexMap = [
        [/#5e72e4\b/gi, 'var(--brand1)'],
        [/#825ee4\b/gi, 'var(--brand2)'],
        [/#667eea\b/gi, 'var(--brand1)'],
        [/#764ba2\b/gi, 'var(--brand2)'],
        [/#6f5bd8\b/gi, 'var(--brand2)'],
        [/#4c3f9a\b/gi, 'color-mix(in srgb, var(--brand1) 55%, #0c1222)'],
    ];
    for (const [re, replacement] of hexMap) {
        content = content.replace(re, replacement);
    }
    return content;
}

function protectAppCssBlocks(content) {
    const placeholders = [];
    let idx = 0;

    const save = (match) => {
        const key = `__PROTECTED_${idx++}__`;
        placeholders.push({ key, value: match });
        return key;
    };

    // Classic brand variable overrides (light + dark)
    content = content.replace(
        /\/\* Klassisch: Violett[\s\S]*?html\[data-brand="classic"\]\[data-theme="dark"\] \{[\s\S]*?\}\n\n/,
        save
    );

    // Intentional purple swatch preview
    content = content.replace(
        /\.ms365-auth-menu__brand-swatch--classic \{[\s\S]*?\}\n/,
        save
    );

    return { content, placeholders };
}

function restorePlaceholders(content, placeholders) {
    for (const { key, value } of placeholders) {
        content = content.replace(key, value);
    }
    return content;
}

function transformContent(content, filePath) {
    const rel = path.relative(ROOT, filePath).replace(/\\/g, '/');

    if (rel === 'app.css') {
        const { content: protectedContent, placeholders } = protectAppCssBlocks(content);
        let transformed = replaceRgba(protectedContent);
        transformed = replaceHex(transformed);
        return restorePlaceholders(transformed, placeholders);
    }

    let transformed = replaceRgba(content);
    transformed = replaceHex(transformed);
    return transformed;
}

const files = walkDir(ROOT);
const changed = [];

for (const file of files) {
    const original = fs.readFileSync(file, 'utf8');
    const transformed = transformContent(original, file);
    if (transformed !== original) {
        fs.writeFileSync(file, transformed, 'utf8');
        changed.push(path.relative(ROOT, file).replace(/\\/g, '/'));
    }
}

console.log(`Changed ${changed.length} file(s):`);
for (const f of changed.sort()) {
    console.log(`  ${f}`);
}
