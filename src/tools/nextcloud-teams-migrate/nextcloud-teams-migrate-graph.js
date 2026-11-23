/**
 * Graph-Helfer: Team-Drive, Kanäle, Upload (Browser).
 * Transport: shared/graph-client.js (R22 / b11).
 */

import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';

const SMALL_MAX = 4 * 1024 * 1024;

/**
 * @param {string} method
 * @param {string} path
 * @param {string} token
 * @param {any} [body]
 * @param {Record<string, string>} [extraHeaders]
 */
export async function graphJson(method, path, token, body, extraHeaders = {}) {
    const rawBody =
        body instanceof Blob ||
        body instanceof ArrayBuffer ||
        ArrayBuffer.isView(body) ||
        typeof body === 'string';
    try {
        return await sharedGraphJson(method, path, token, body === null ? undefined : body, extraHeaders, {
            rawBody: rawBody && body !== undefined && body !== null
        });
    } catch (e) {
        const status = /** @type {any} */ (e).status || 0;
        const msg = String((e && e.message) || e);
        const err = new Error(`Graph ${method} ${path} → ${status}: ${msg}`);
        /** @type {any} */ (err).status = status;
        throw err;
    }
}

/**
 * @param {string} token
 * @param {string} groupId
 */
export async function getGroupDrive(token, groupId) {
    const drive = await graphJson('GET', `/groups/${encodeURIComponent(groupId)}/drive`, token);
    if (!drive || !drive.id) throw new Error('Keine Dokumentbibliothek für dieses Team.');
    return drive;
}

/**
 * @param {string} token
 * @param {string} teamId
 */
export async function listChannels(token, teamId) {
    const data = await graphJson(
        'GET',
        `/teams/${encodeURIComponent(teamId)}/channels?$select=id,displayName,membershipType`,
        token
    );
    return (data && data.value) || [];
}

/**
 * Öffentlichen Kanal anlegen oder vorhandene ID zurückgeben.
 * @param {string} token
 * @param {string} teamId
 * @param {string} displayName
 * @param {Array<{ id: string, displayName: string }>} [existing]
 */
export async function ensureChannel(token, teamId, displayName, existing) {
    const name = String(displayName || '').trim().slice(0, 50);
    if (!name) throw new Error('Kanalname fehlt');
    const list = existing || (await listChannels(token, teamId));
    const hit = list.find(
        (c) => String(c.displayName || '').toLowerCase() === name.toLowerCase()
    );
    if (hit) return hit;
    const created = await graphJson(
        'POST',
        `/teams/${encodeURIComponent(teamId)}/channels`,
        token,
        {
            displayName: name,
            description: 'Import aus Nextcloud',
            membershipType: 'standard'
        }
    );
    return created;
}

/**
 * Dateien-Ordner eines Kanals (SharePoint-Unterordner).
 * @param {string} token
 * @param {string} teamId
 * @param {string} channelId
 */
export async function getChannelFilesFolder(token, teamId, channelId) {
    const folder = await graphJson(
        'GET',
        `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/filesFolder`,
        token
    );
    if (!folder || !folder.id || !folder.parentReference || !folder.parentReference.driveId) {
        throw new Error('Kanal-Dateienordner nicht gefunden (filesFolder).');
    }
    return {
        driveId: folder.parentReference.driveId,
        folderId: folder.id,
        name: folder.name || '',
        webUrl: folder.webUrl || ''
    };
}

/**
 * @param {string} token
 * @param {string} driveId
 * @param {string} relativeFolderPath
 */
export async function ensureFolderPath(token, driveId, relativeFolderPath) {
    const parts = String(relativeFolderPath || '')
        .replace(/^\/+|\/+$/g, '')
        .split(/[\\/]/)
        .map((p) => p.trim())
        .filter(Boolean);
    if (!parts.length) {
        return graphJson('GET', `/drives/${encodeURIComponent(driveId)}/root`, token);
    }
    /** @type {any} */
    let last = null;
    /** @type {string[]} */
    const built = [];
    for (const part of parts) {
        built.push(part);
        const encodedPath = built.map((p) => encodeURIComponent(p)).join('/');
        const getUrl = `/drives/${encodeURIComponent(driveId)}/root:/${encodedPath}`;
        try {
            last = await graphJson('GET', getUrl, token);
            continue;
        } catch (e) {
            if (/** @type {any} */ (e).status !== 404) throw e;
        }
        const parentId = (last && last.id) || 'root';
        try {
            last = await graphJson(
                'POST',
                `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children`,
                token,
                {
                    name: part,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'fail'
                }
            );
        } catch (e) {
            if (/** @type {any} */ (e).status === 409) {
                last = await graphJson('GET', getUrl, token);
            } else {
                throw e;
            }
        }
    }
    return last;
}

/**
 * Unterordner relativ zu einem bekannten Parent-Item.
 * @param {string} token
 * @param {string} driveId
 * @param {string} parentFolderId
 * @param {string} relativeFolderPath
 */
export async function ensureChildFolderPath(token, driveId, parentFolderId, relativeFolderPath) {
    const parts = String(relativeFolderPath || '')
        .replace(/^\/+|\/+$/g, '')
        .split(/[\\/]/)
        .map((p) => p.trim())
        .filter(Boolean);
    if (!parts.length) {
        return { id: parentFolderId };
    }
    let parentId = parentFolderId;
    /** @type {any} */
    let last = { id: parentFolderId };
    for (const part of parts) {
        const children = await graphJson(
            'GET',
            `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children?$select=id,name,folder&$top=200`,
            token
        );
        const found = ((children && children.value) || []).find(
            (c) => c.folder && String(c.name || '') === part
        );
        if (found) {
            last = found;
            parentId = found.id;
            continue;
        }
        try {
            last = await graphJson(
                'POST',
                `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children`,
                token,
                {
                    name: part,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'fail'
                }
            );
        } catch (e) {
            if (/** @type {any} */ (e).status === 409) {
                const again = await graphJson(
                    'GET',
                    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children?$select=id,name,folder&$top=200`,
                    token
                );
                last = ((again && again.value) || []).find(
                    (c) => c.folder && String(c.name || '') === part
                );
                if (!last) throw e;
            } else {
                throw e;
            }
        }
        parentId = last.id;
    }
    return last;
}

/**
 * @param {string} token
 * @param {string} driveId
 * @param {string} parentFolderId
 * @param {string} fileName
 * @param {Blob} blob
 */
export async function uploadBlob(token, driveId, parentFolderId, fileName, blob) {
    const name = String(fileName || '').trim();
    if (!name) throw new Error('Dateiname leer');
    const size = blob.size || 0;

    if (size < SMALL_MAX) {
        return graphJson(
            'PUT',
            `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentFolderId)}:/${encodeURIComponent(name)}:/content`,
            token,
            blob,
            { 'Content-Type': blob.type || 'application/octet-stream' }
        );
    }

    const session = await graphJson(
        'POST',
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentFolderId)}:/${encodeURIComponent(name)}:/createUploadSession`,
        token,
        {
            item: {
                '@microsoft.graph.conflictBehavior': 'replace',
                name
            }
        }
    );
    const uploadUrl = session && session.uploadUrl;
    if (!uploadUrl) throw new Error('createUploadSession ohne uploadUrl');

    const chunkSize = 5 * 1024 * 1024;
    let offset = 0;
    /** @type {any} */
    let last = null;
    while (offset < size) {
        const end = Math.min(offset + chunkSize, size);
        const chunk = blob.slice(offset, end);
        const res = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Content-Length': String(chunk.size),
                'Content-Range': `bytes ${offset}-${end - 1}/${size}`
            },
            body: chunk
        });
        const text = await res.text();
        let data = null;
        try {
            data = text ? JSON.parse(text) : null;
        } catch {
            data = { raw: text };
        }
        if (!res.ok && res.status !== 202) {
            throw new Error(`Upload-Chunk ${res.status}: ${text.slice(0, 240)}`);
        }
        last = data;
        offset = end;
    }
    return last;
}

/**
 * @template T
 * @param {() => Promise<T>} fn
 * @param {{ retries?: number }} [opts]
 */
export async function withRetry(fn, opts = {}) {
    const retries = opts.retries ?? 3;
    let lastErr;
    for (let i = 0; i <= retries; i++) {
        try {
            return await fn();
        } catch (e) {
            lastErr = e;
            const status = /** @type {any} */ (e).status || 0;
            if (!(status === 429 || status >= 500 || status === 0) || i === retries) break;
            await new Promise((r) => setTimeout(r, Math.min(30000, 1000 * Math.pow(2, i))));
        }
    }
    throw lastErr;
}
