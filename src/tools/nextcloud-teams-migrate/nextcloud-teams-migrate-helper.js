/**
 * Client für den lokalen Nextcloud-Helfer (proxy-server.mjs auf 127.0.0.1).
 */
import { norm } from './nextcloud-teams-migrate-logic.js';
import { parsePropfindXml } from './nextcloud-teams-migrate-webdav.js';

export const DEFAULT_HELPER_BASE = 'http://127.0.0.1:3791';

/**
 * @param {string} [base]
 */
export function helperBase(base) {
    return norm(base) || DEFAULT_HELPER_BASE;
}

/**
 * @param {string} base
 * @param {string} relPath
 */
function davPath(base, relPath) {
    const root = helperBase(base).replace(/\/+$/, '') + '/dav';
    const parts = norm(relPath)
        .replace(/^\/+|\/+$/g, '')
        .split('/')
        .filter(Boolean)
        .map((p) => encodeURIComponent(p));
    return parts.length ? root + '/' + parts.join('/') : root;
}

/**
 * @param {string} [base]
 */
export async function helperHealth(base) {
    const url = helperBase(base).replace(/\/+$/, '') + '/health';
    let res;
    try {
        res = await fetch(url);
    } catch (e) {
        const err = new Error(
            'Lokaler Helfer läuft nicht. Bitte „Start-Helfer.cmd“ doppelklicken und das Fenster offen lassen. (' +
                (e && e.message ? e.message : e) +
                ')'
        );
        /** @type {any} */ (err).code = 'HELPER_DOWN';
        throw err;
    }
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) {
        const err = new Error(
            (data && data.error) ||
                'Helfer erreichbar, aber Nextcloud-Login fehlgeschlagen. .env.local prüfen (App-Passwort).'
        );
        /** @type {any} */ (err).code = 'HELPER_AUTH';
        /** @type {any} */ (err).data = data;
        throw err;
    }
    return data;
}

/**
 * @param {{ base?: string, relPath?: string, depth?: string, userForParse?: string }} opts
 */
export async function helperPropfind(opts) {
    const base = helperBase(opts.base);
    const url = davPath(base, opts.relPath || '');
    const res = await fetch(url, {
        method: 'PROPFIND',
        headers: {
            Depth: opts.depth || '1',
            'Content-Type': 'application/xml; charset=utf-8'
        },
        body:
            '<?xml version="1.0"?>' +
            '<d:propfind xmlns:d="DAV:">' +
            '<d:prop><d:resourcetype/><d:getcontentlength/><d:displayname/></d:prop>' +
            '</d:propfind>'
    });
    const text = await res.text();
    if (!res.ok) {
        throw new Error(`Helfer PROPFIND ${res.status}: ${text.slice(0, 240)}`);
    }
    // user nur für href-Parsing; Helfer liefert Nextcloud-XML mit echtem User in href
    const user = norm(opts.userForParse) || 'user';
    return parsePropfindXml(text, user);
}

/**
 * @param {{ base?: string, relPath?: string, userForParse?: string }} opts
 */
export async function helperListChildren(opts) {
    const rel = norm(opts.relPath).replace(/^\/+|\/+$/g, '');
    const items = await helperPropfind({ ...opts, relPath: rel, depth: '1' });
    return items
        .filter((it) => it.relPath !== rel)
        .sort((a, b) => {
            if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
            return a.name.localeCompare(b.name, 'de', { sensitivity: 'base' });
        });
}

/**
 * @param {{ base?: string, relPath: string, onProgress?: Function, userForParse?: string }} opts
 */
export async function helperInventoryFiles(opts) {
    const baseRel = norm(opts.relPath).replace(/^\/+|\/+$/g, '');
    /** @type {Array<{ relPath: string, name: string, size: number }>} */
    const files = [];
    /** @type {string[]} */
    const queue = [baseRel];
    const seen = new Set();
    let folderCount = 0;
    while (queue.length) {
        const current = queue.shift() || '';
        if (seen.has(current)) continue;
        seen.add(current);
        const children = await helperListChildren({
            base: opts.base,
            relPath: current,
            userForParse: opts.userForParse
        });
        for (const it of children) {
            if (it.isFolder) {
                folderCount++;
                queue.push(it.relPath);
            } else {
                files.push({ relPath: it.relPath, name: it.name, size: it.size || 0 });
            }
        }
        if (typeof opts.onProgress === 'function') {
            opts.onProgress({
                folders: folderCount,
                files: files.length,
                bytes: files.reduce((s, f) => s + (f.size || 0), 0)
            });
        }
    }
    files.sort((a, b) => a.relPath.localeCompare(b.relPath, 'de'));
    return {
        baseRel,
        files,
        folderCount,
        fileCount: files.length,
        totalBytes: files.reduce((s, f) => s + (f.size || 0), 0)
    };
}

/**
 * @param {{ base?: string, relPath: string }} opts
 */
export async function helperDownloadBlob(opts) {
    const url = davPath(opts.base, opts.relPath);
    const res = await fetch(url, { method: 'GET' });
    if (!res.ok) {
        const t = await res.text().catch(() => '');
        throw new Error(`Helfer GET ${res.status}: ${t.slice(0, 200)}`);
    }
    return res.blob();
}
