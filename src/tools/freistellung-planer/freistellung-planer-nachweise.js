/**
 * Nachweise (Anhänge) für Freistellungsanträge – Ablage in Site Assets / Standard-Bibliothek.
 */
import { resolveConfigDrive, encodeDriveRootPathForUpload } from './freistellung-planer-remote-config.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

export const NACHWEISE_MAX_FILES = 5;
export const NACHWEISE_MAX_BYTES = 8 * 1024 * 1024;

export function formatNachweisSize(bytes) {
    const n = Number(bytes) || 0;
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return Math.round(n / 1024) + ' KB';
    return (n / (1024 * 1024)).toFixed(1) + ' MB';
}

const FOLDER = 'ms365/Freistellungsnachweise';

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

function sanitizeFileName(name) {
    return String(name || 'datei')
        .replace(/[<>:"/\\|?*]/g, '_')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 180);
}

/**
 * @param {unknown} raw
 * @returns {Array<{ name: string, url: string }>}
 */
export function parseNachweiseField(raw) {
    const text = String(raw || '').trim();
    if (!text) return [];
    try {
        const j = JSON.parse(text);
        if (!Array.isArray(j)) return [];
        return j
            .map((x) => ({
                name: String((x && x.name) || '').trim(),
                url: String((x && x.url) || '').trim()
            }))
            .filter((x) => x.name && x.url);
    } catch {
        return [];
    }
}

/**
 * @param {Array<{ name: string, url: string }>} links
 */
export function stringifyNachweiseField(links) {
    const clean = (links || [])
        .map((x) => ({
            name: String(x.name || '').trim(),
            url: String(x.url || '').trim()
        }))
        .filter((x) => x.name && x.url);
    return JSON.stringify(clean);
}

/**
 * @param {FileList|File[]} fileList
 */
export function readPendingUploadFiles(fileList) {
    const files = fileList ? Array.from(fileList) : [];
    if (files.length > NACHWEISE_MAX_FILES) {
        throw new Error('Maximal ' + NACHWEISE_MAX_FILES + ' Dateien pro Antrag.');
    }
    files.forEach((f) => {
        if (f.size > NACHWEISE_MAX_BYTES) {
            throw new Error('Datei zu groß: ' + f.name + ' (max. ' + Math.round(NACHWEISE_MAX_BYTES / 1024 / 1024) + ' MB).');
        }
    });
    return files;
}

/**
 * @param {{ siteId: string, webUrl?: string }} ctx
 * @param {string} itemId
 * @param {File[]} files
 * @returns {Promise<Array<{ name: string, url: string }>>}
 */
export async function uploadFreistellungNachweise(ctx, itemId, files) {
    const id = String(itemId || '').trim();
    if (!id || !files || !files.length) return [];
    const webUrl = String((ctx && ctx.webUrl) || '').trim();
    if (!webUrl) throw new Error('Site-URL für Nachweise fehlt.');
    const tok = await G().getGraphToken(SCOPES);
    const { driveId } = await resolveConfigDrive(tok, webUrl);
    const uploaded = [];
    for (const file of files) {
        const rel = FOLDER + '/' + id + '/' + sanitizeFileName(file.name);
        const enc = encodeDriveRootPathForUpload(rel);
        const putUrl = G().graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/' + enc + '/content';
        const res = await fetch(putUrl, {
            method: 'PUT',
            headers: {
                Authorization: 'Bearer ' + tok,
                'Content-Type': file.type || 'application/octet-stream'
            },
            body: file
        });
        if (!res.ok) {
            const text = await res.text();
            throw new Error('Upload fehlgeschlagen (' + file.name + '): ' + (text || res.status));
        }
        let meta = null;
        try {
            meta = await res.json();
        } catch {
            meta = null;
        }
        const url =
            (meta && meta.webUrl) ||
            (meta && meta['@microsoft.graph.downloadUrl']) ||
            '';
        uploaded.push({
            name: file.name,
            url: url || putUrl
        });
    }
    return uploaded;
}

/**
 * @param {{ siteId: string, list: { id: string } }} ctx
 * @param {string} itemId
 * @param {Array<{ name: string, url: string }>} links
 */
export async function saveNachweiseOnItem(ctx, itemId, links) {
    const tok = await G().getGraphToken(SCOPES);
    await G().graphJson(
        'PATCH',
        G().graphPathSite(ctx.siteId) +
            '/lists/' +
            encodeURIComponent(ctx.list.id) +
            '/items/' +
            encodeURIComponent(itemId) +
            '/fields',
        tok,
        { Nachweise: stringifyNachweiseField(links) },
        'v1.0'
    );
}
