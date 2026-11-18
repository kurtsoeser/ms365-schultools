/**
 * Nextcloud WebDAV: PROPFIND + GET (stream).
 */
import { Readable } from 'node:stream';

/**
 * @param {string} baseUrl
 * @param {string} user
 */
export function davFilesRoot(baseUrl, user) {
    const base = String(baseUrl || '').replace(/\/+$/, '');
    const u = encodeURIComponent(String(user || '').trim());
    return `${base}/remote.php/dav/files/${u}`;
}

/**
 * @param {string} root
 * @param {string} relPath
 */
export function davUrl(root, relPath) {
    const parts = String(relPath || '')
        .replace(/^\/+|\/+$/g, '')
        .split('/')
        .filter(Boolean)
        .map((p) => encodeURIComponent(p));
    return parts.length ? `${root}/${parts.join('/')}` : root;
}

/**
 * @param {string} user
 * @param {string} appPassword
 */
export function basicAuthHeader(user, appPassword) {
    const token = Buffer.from(`${user}:${appPassword}`, 'utf8').toString('base64');
    return `Basic ${token}`;
}

/**
 * Minimal XML unescape.
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
 * @param {string} filesRootHref decoded path prefix ending without trailing slash logic
 */
function hrefToRelPath(href, filesRootPath) {
    let pathPart = String(href || '');
    try {
        if (/^https?:\/\//i.test(pathPart)) {
            pathPart = new URL(pathPart).pathname;
        }
    } catch {
        /* ignore */
    }
    let decoded = pathPart;
    try {
        decoded = decodeURIComponent(pathPart);
    } catch {
        decoded = pathPart;
    }
    const root = filesRootPath.endsWith('/') ? filesRootPath.slice(0, -1) : filesRootPath;
    // auch encoded Username in href abfangen
    const encRoot = root.replace(/\/([^/]+)$/, (_, u) => '/' + encodeURIComponent(u));
    let rel = decoded;
    for (const marker of [root, encRoot]) {
        const idx = decoded.toLowerCase().indexOf(marker.toLowerCase());
        if (idx >= 0) {
            rel = decoded.slice(idx + marker.length);
            break;
        }
        try {
            const idx2 = pathPart.toLowerCase().indexOf(marker.toLowerCase());
            if (idx2 >= 0) {
                rel = decodeURIComponent(pathPart.slice(idx2 + marker.length));
                break;
            }
        } catch {
            /* ignore */
        }
    }
    return String(rel).replace(/^\/+|\/+$/g, '');
}

/**
 * @typedef {{ href: string, relPath: string, isCollection: boolean, size: number, contentType: string, displayName: string }} DavItem
 */

/**
 * @param {string} xml
 * @param {string} filesRootPath e.g. /remote.php/dav/files/user
 * @returns {DavItem[]}
 */
export function parsePropfindXml(xml, filesRootPath) {
    const responses = String(xml || '').split(/<d:response\b/i).slice(1);
    /** @type {DavItem[]} */
    const items = [];
    for (const chunk of responses) {
        const hrefMatch = chunk.match(/<d:href>([^<]*)<\/d:href>/i);
        if (!hrefMatch) continue;
        const href = xmlUnescape(hrefMatch[1]);
        const isCollection = /<d:collection\s*\/>/i.test(chunk) || /<d:resourcetype[^>]*>\s*<d:collection/i.test(chunk);
        const sizeMatch = chunk.match(/<d:getcontentlength>([^<]*)<\/d:getcontentlength>/i);
        const typeMatch = chunk.match(/<d:getcontenttype>([^<]*)<\/d:getcontenttype>/i);
        const nameMatch = chunk.match(/<d:displayname>([^<]*)<\/d:displayname>/i);
        const relPath = hrefToRelPath(href, filesRootPath);
        items.push({
            href,
            relPath,
            isCollection,
            size: sizeMatch ? Number(sizeMatch[1]) || 0 : 0,
            contentType: typeMatch ? xmlUnescape(typeMatch[1]) : '',
            displayName: nameMatch ? xmlUnescape(nameMatch[1]) : ''
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
 *   depth?: '0' | '1' | 'infinity'
 * }} opts
 */
export async function propfind(opts) {
    const root = davFilesRoot(opts.baseUrl, opts.user);
    const url = davUrl(root, opts.relPath || '');
    const depth = opts.depth || '1';
    const filesRootPath = `/remote.php/dav/files/${opts.user}`;

    const res = await fetch(url, {
        method: 'PROPFIND',
        headers: {
            Authorization: basicAuthHeader(opts.user, opts.appPassword),
            Depth: depth,
            'Content-Type': 'application/xml; charset=utf-8'
        },
        body:
            '<?xml version="1.0"?>' +
            '<d:propfind xmlns:d="DAV:">' +
            '<d:prop>' +
            '<d:resourcetype/>' +
            '<d:getcontentlength/>' +
            '<d:getcontenttype/>' +
            '<d:displayname/>' +
            '</d:prop>' +
            '</d:propfind>'
    });

    const text = await res.text();
    if (!res.ok) {
        throw new Error(`WebDAV PROPFIND ${res.status} ${url}: ${text.slice(0, 300)}`);
    }
    return parsePropfindXml(text, filesRootPath);
}

/**
 * Rekursives Inventar (Depth:1 Walk – robust, falls infinity blockiert).
 * @param {{
 *   baseUrl: string,
 *   user: string,
 *   appPassword: string,
 *   relPath: string
 * }} opts
 */
export async function inventoryTree(opts) {
    const baseRel = String(opts.relPath || '').replace(/^\/+|\/+$/g, '');
    /** @type {DavItem[]} */
    const files = [];
    /** @type {DavItem[]} */
    const folders = [];
    let totalBytes = 0;

    /** @type {string[]} */
    const queue = [baseRel];
    const seen = new Set();

    while (queue.length) {
        const current = queue.shift() || '';
        if (seen.has(current)) continue;
        seen.add(current);

        let items;
        try {
            items = await propfind({
                baseUrl: opts.baseUrl,
                user: opts.user,
                appPassword: opts.appPassword,
                relPath: current,
                depth: '1'
            });
        } catch (e) {
            throw new Error(
                `Inventar fehlgeschlagen bei „${current || '/'}“: ${e && e.message ? e.message : e}`
            );
        }

        for (const it of items) {
            if (!it.relPath && !current) {
                // root response
                if (it.isCollection) folders.push(it);
                continue;
            }
            if (it.relPath === current) {
                if (it.isCollection) folders.push(it);
                continue;
            }
            if (it.isCollection) {
                folders.push(it);
                queue.push(it.relPath);
            } else {
                files.push(it);
                totalBytes += it.size || 0;
            }
        }
    }

    files.sort((a, b) => a.relPath.localeCompare(b.relPath, 'de'));
    folders.sort((a, b) => a.relPath.localeCompare(b.relPath, 'de'));

    return {
        baseRel,
        files,
        folders,
        fileCount: files.length,
        folderCount: folders.length,
        totalBytes
    };
}

/**
 * Datei als Node Readable streamen.
 * @param {{
 *   baseUrl: string,
 *   user: string,
 *   appPassword: string,
 *   relPath: string
 * }} opts
 */
export async function downloadFileStream(opts) {
    const root = davFilesRoot(opts.baseUrl, opts.user);
    const url = davUrl(root, opts.relPath);
    const res = await fetch(url, {
        method: 'GET',
        headers: {
            Authorization: basicAuthHeader(opts.user, opts.appPassword)
        }
    });
    if (!res.ok) {
        const t = await res.text().catch(() => '');
        throw new Error(`WebDAV GET ${res.status} ${url}: ${t.slice(0, 200)}`);
    }
    if (!res.body) throw new Error('WebDAV GET ohne Body: ' + url);
    return {
        url,
        contentType: res.headers.get('content-type') || 'application/octet-stream',
        contentLength: Number(res.headers.get('content-length') || 0) || null,
        stream: Readable.fromWeb(/** @type {any} */ (res.body))
    };
}

/**
 * @param {number} n
 */
export function formatBytes(n) {
    const v = Number(n) || 0;
    if (v < 1024) return `${v} B`;
    if (v < 1024 * 1024) return `${(v / 1024).toFixed(1)} KB`;
    if (v < 1024 * 1024 * 1024) return `${(v / (1024 * 1024)).toFixed(2)} MB`;
    return `${(v / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}
