/**
 * Kopiert veröffentlichte Neuigkeiten auf die Homepage.
 * Quelle: public/release-notes.json und public/release-notes/
 * Ziel:   landing/release-notes.json und landing/release-notes/
 */
import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const srcJson = path.join(root, 'public', 'release-notes.json');
const srcDir = path.join(root, 'public', 'release-notes');
const destJson = path.join(root, 'landing', 'release-notes.json');
const destDir = path.join(root, 'landing', 'release-notes');

if (!fs.existsSync(srcJson)) {
    console.warn('public/release-notes.json fehlt, Homepage-Kopie übersprungen.');
    process.exit(0);
}

fs.copyFileSync(srcJson, destJson);
fs.mkdirSync(destDir, { recursive: true });

if (fs.existsSync(srcDir)) {
    for (const name of fs.readdirSync(srcDir)) {
        const from = path.join(srcDir, name);
        if (!fs.statSync(from).isFile()) continue;
        fs.copyFileSync(from, path.join(destDir, name));
    }
}

console.log('Neuigkeiten nach landing/ kopiert.');
