/**
 * Browser-WebDAV-Client für Nextcloud (PROPFIND/GET).
 * Hinweis: Zielserver muss CORS für die App-Origin erlauben (oder CORS-Proxy).
 */
import { normalizeNextcloudBase, norm } from './nextcloud-teams-migrate-logic.js';

/**
 * @param {string} baseUrl
 * @param {string} user
 * @param {string} [corsProxy]
 */
export function davFilesRoot(baseUrl, user, corsProxy) {
    const base = normalizeNextcloudBase(baseUrl);
    const u = encodeURIComponent(norm(user));
    const url = `${base}/remote.php/dav/files/${u}`;
    const proxy = norm(corsProxy);
    if (!proxy) return url;
    // Proxy erwartet typischerweise die Ziel-URL angehängt
    return proxy.replace(/\/+$/, '') + '/' + url;
}

/**
 * @param {string} root
 * @param {string} relPath
 */
export function davUrl(root, relPath) {
    const parts = norm(relPath)
        .replace(/^\/+|\/+$/g, '')
        .split('/')
        .filter(Boolean)
        .map((p) => encodeURIComponent(p));
    // Wenn Proxy den vollen Root schon enthält: Segmente anhängen
    if (!parts.length) return root;
    // Proxy-Form: proxy/https://host/.../user  → Segmente mit / anhängen
    return `${root}/${parts.join('/')}`;
}

/**
 * @param {string} user
 * @param {string} appPassword
 */
export function basicAuthHeader(user, appPassword) {
    // btoa erwartet Latin1 – für typische User/App-Passwörter ok
    return 'Basic ' + btoa(`${user}:${appPassword}`);
}

/**
 * @param {string} s
 */
function xmlUnescape(s) {
    return String(s || '')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
}

/**
 * @param {string} href
 * @param {string} user
 */
function hrefToRelPath(href, user) {
    let pathPart = String(href || '');
    try {
        if (/^https?:\/\//i.test(pathPart)) pathPart = new URL(pathPart).pathname;
    } catch {
        /* ignore */
    }
    let decoded = pathPart;
    try {
        decoded = decodeURIComponent(pathPart);
    } catch {
        decoded = pathPart;
    }
    const roots = [
        `/remote.php/dav/files/${user}`,
        `/remote.php/dav/files/${encodeURIComponent(user)}`
    ];
    for (const root of roots) {
        const idx = decoded.toLowerCase().indexOf(root.toLowerCase());
        if (idx >= 0) return decoded.slice(idx + root.length).replace(/^\/+|\/+$/g, '');
        try {
            const idx2 = pathPart.toLowerCase().indexOf(root.toLowerCase());
            if (idx2 >= 0) {
                return decodeURIComponent(pathPart.slice(idx2 + root.length)).replace(/^\/+|\/+$/g, '');
            }
        } catch {
            /* ignore */
        }
    }
    return decoded.replace(/^\/+|\/+$/g, '');
}

/**
 * @param {string} xml
 * @param {string} user
 */
export function parsePropfindXml(xml, user) {
    const responses = String(xml || '').split(/<d:response\b/i).slice(1);
    /** @type {Array<{ href: string, relPath: string, isFolder: boolean, size: number, name: string }>} */
    const items = [];
    for (const chunk of responses) {
        const hrefMatch = chunk.match(/<d:href>([^<]*)<\/d:href>/i);
        if (!hrefMatch) continue;
        const href = xmlUnescape(hrefMatch[1]);
        const isFolder =
            /<d:collection\s*\/>/i.test(chunk) || /<d:resourcetype[^>]*>\s*<d:collection/i.test(chunk);
        const sizeMatch = chunk.match(/<d:getcontentlength>([^<]*)<\/d:getcontentlength>/i);
        const nameMatch = chunk.match(/<d:displayname>([^<]*)<\/d:displayname>/i);
        const relPath = hrefToRelPath(href, user);
        const name =
            (nameMatch ? xmlUnescape(nameMatch[1]) : '') ||
            (relPath.split('/').filter(Boolean).pop() || '');
        items.push({
            href,
            relPath,
            isFolder,
            size: sizeMatch ? Number(sizeMatch[1]) || 0 : 0,
            name
        });
    }
    return items;
}

/**
 * @param {{
 *   baseUrl: string,
 *   user: string,
 *   appPassword: string,
 *   relPath?: string,
 *   depth?: '0'|'1',
 *   corsProxy?: string
 * }} opts
 */
export async function propfind(opts) {
    const root = davFilesRoot(opts.baseUrl, opts.user, opts.corsProxy);
    const url = davUrl(root, opts.relPath || '');
    let res;
    try {
        res = await fetch(url, {
            method: 'PROPFIND',
            headers: {
                Authorization: basicAuthHeader(opts.user, opts.appPassword),
                Depth: opts.depth || '1',
                'Content-Type': 'application/xml; charset=utf-8'
            },
            body:
                '<?xml version="1.0"?>' +
                '<d:propfind xmlns:d="DAV:">' +
                '<d:prop><d:resourcetype/><d:getcontentlength/><d:displayname/></d:prop>' +
                '</d:propfind>'
        });
    } catch (e) {
        const err = new Error(
            'Nextcloud nicht erreichbar (oft CORS). Siehe Hinweis im Tool oder CLI unter scripts/nextcloud-teams-migrate/. ' +
                (e && e.message ? e.message : e)
        );
        /** @type {any} */ (err).code = 'CORS_OR_NETWORK';
        throw err;
    }
    const text = await res.text();
    if (!res.ok) {
        throw new Error(`WebDAV PROPFIND ${res.status}: ${text.slice(0, 240)}`);
    }
    return parsePropfindXml(text, opts.user);
}

/**
 * Kinder eines Ordners (ohne Self).
 * @param {Parameters<typeof propfind>[0]} opts
 */
export async function listChildren(opts) {
    const rel = norm(opts.relPath).replace(/^\/+|\/+$/g, '');
    const items = await propfind({ ...opts, relPath: rel, depth: '1' });
    return items
        .filter((it) => it.relPath !== rel)
        .sort((a, b) => {
            if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
            return a.name.localeCompare(b.name, 'de', { sensitivity: 'base' });
        });
}

/**
 * Rekursives Datei-Inventar unter relPath.
 * @param {Parameters<typeof propfind>[0] & { onProgress?: (info: { folders: number, files: number, bytes: number }) => void }} opts
 */
export async function inventoryFiles(opts) {
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
        const children = await listChildren({ ...opts, relPath: current });
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
 * @param {{
 *   baseUrl: string,
 *   user: string,
 *   appPassword: string,
 *   relPath: string,
 *   corsProxy?: string
 * }} opts
 * @returns {Promise<Blob>}
 */
export async function downloadBlob(opts) {
    const root = davFilesRoot(opts.baseUrl, opts.user, opts.corsProxy);
    const url = davUrl(root, opts.relPath);
    let res;
    try {
        res = await fetch(url, {
            method: 'GET',
            headers: { Authorization: basicAuthHeader(opts.user, opts.appPassword) }
        });
    } catch (e) {
        const err = new Error(
            'Download fehlgeschlagen (CORS/Netzwerk): ' + (e && e.message ? e.message : e)
        );
        /** @type {any} */ (err).code = 'CORS_OR_NETWORK';
        throw err;
    }
    if (!res.ok) {
        const t = await res.text().catch(() => '');
        throw new Error(`WebDAV GET ${res.status}: ${t.slice(0, 200)}`);
    }
    return res.blob();
}
