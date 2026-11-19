/**
 * Kopiert die Browser-Bibliotheken aus node_modules nach vendor/ (gitignored).
 *
 * Hintergrund: Die Seiten laden viele Skripte ungebündelt aus src/ (klassische
 * <script>-Tags, dist/src/ ist eine 1:1-Kopie). Ein `import '@azure/msal-browser'`
 * würde dort im Browser scheitern. Deshalb liefern wir die fertigen
 * Browser-Builds der npm-Pakete unter derselben Origin aus, statt sie zur
 * Laufzeit von fremden CDNs zu laden. Versionen und Integrität sichert
 * package-lock.json.
 *
 *   node scripts/vendor-libs.mjs
 *
 * Läuft automatisch vor `npm run dev` und `npm run build`.
 */
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const modulesRoot = path.join(projectRoot, 'node_modules');
const vendorRoot = path.join(projectRoot, 'vendor');

/**
 * `files`: Quelle (relativ zum Paket) → Ziel (relativ zu vendor/<dir>/).
 * Die Zielpfade sind in HTML und src/shared/vendor-libs.js fest verdrahtet.
 */
export const VENDOR_LIBS = [
    {
        pkg: '@azure/msal-browser',
        dir: 'msal-browser',
        files: { 'lib/msal-browser.min.js': 'msal-browser.min.js' }
    },
    {
        pkg: 'xlsx',
        dir: 'xlsx',
        files: { 'dist/xlsx.full.min.js': 'xlsx.full.min.js' }
    },
    {
        pkg: 'jszip',
        dir: 'jszip',
        files: { 'dist/jszip.min.js': 'jszip.min.js' }
    },
    {
        pkg: 'pdfjs-dist',
        dir: 'pdfjs',
        files: {
            'build/pdf.min.js': 'pdf.min.js',
            'build/pdf.worker.min.js': 'pdf.worker.min.js'
        }
    },
    {
        pkg: 'mermaid',
        dir: 'mermaid',
        files: { 'dist/mermaid.min.js': 'mermaid.min.js' }
    },
    {
        pkg: 'bootstrap-icons',
        dir: 'bootstrap-icons',
        // Das CSS verweist relativ auf ./fonts/ – Struktur beibehalten.
        files: {
            'font/bootstrap-icons.min.css': 'bootstrap-icons.min.css',
            'font/fonts/bootstrap-icons.woff2': 'fonts/bootstrap-icons.woff2',
            'font/fonts/bootstrap-icons.woff': 'fonts/bootstrap-icons.woff'
        }
    }
];

const LICENSE_NAMES = ['LICENSE', 'LICENSE.md', 'LICENSE.txt', 'LICENSE.markdown', 'LICENCE'];

async function sha256(file) {
    return createHash('sha256').update(await fs.readFile(file)).digest('hex');
}

async function copyLicense(pkgRoot, destDir) {
    for (const name of LICENSE_NAMES) {
        try {
            await fs.copyFile(path.join(pkgRoot, name), path.join(destDir, 'LICENSE.txt'));
            return;
        } catch {
            /* nächster Kandidat */
        }
    }
}

export async function vendorLibs() {
    const manifest = [];
    for (const lib of VENDOR_LIBS) {
        const pkgRoot = path.join(modulesRoot, ...lib.pkg.split('/'));
        let version;
        try {
            version = JSON.parse(await fs.readFile(path.join(pkgRoot, 'package.json'), 'utf8')).version;
        } catch {
            throw new Error(`vendor-libs: Paket "${lib.pkg}" fehlt in node_modules – zuerst "npm ci" ausführen.`);
        }
        const destDir = path.join(vendorRoot, lib.dir);
        const files = {};
        for (const [from, to] of Object.entries(lib.files)) {
            const src = path.join(pkgRoot, from);
            const dst = path.join(destDir, to);
            await fs.mkdir(path.dirname(dst), { recursive: true });
            try {
                await fs.copyFile(src, dst);
            } catch (err) {
                throw new Error(`vendor-libs: ${lib.pkg}@${version} enthält "${from}" nicht (${err.code || err.message}).`);
            }
            files[`${lib.dir}/${to}`] = await sha256(dst);
        }
        await copyLicense(pkgRoot, destDir);
        manifest.push({ package: lib.pkg, version, files });
    }
    await fs.writeFile(
        path.join(vendorRoot, 'vendor-manifest.json'),
        `${JSON.stringify({ libraries: manifest }, null, 2)}\n`,
        'utf8'
    );
    return manifest;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    const manifest = await vendorLibs();
    console.log(`vendor-libs: ${manifest.map((m) => `${m.package}@${m.version}`).join(', ')} → vendor/`);
}
