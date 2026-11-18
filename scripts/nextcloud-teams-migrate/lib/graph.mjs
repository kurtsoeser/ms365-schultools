/**
 * Microsoft Graph: Device-Code Auth + Drive Upload (klein + Upload-Session).
 */
import { Readable } from 'node:stream';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = [
    'offline_access',
    'User.Read',
    'Group.Read.All',
    'Files.ReadWrite.All',
    'Sites.ReadWrite.All'
];

/**
 * @param {number} ms
 */
function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
}

/**
 * @param {string} tenantId
 * @param {string} clientId
 */
export async function acquireDeviceCodeToken(tenantId, clientId) {
    const tenant = String(tenantId || 'organizations').trim() || 'organizations';
    const cid = String(clientId || '').trim();
    if (!cid) {
        throw new Error(
            'GRAPH_CLIENT_ID fehlt (oder ms365-config.js clientId). Entra-App muss „öffentliche Clientflows“ erlauben.'
        );
    }

    const deviceUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/devicecode`;
    const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
    const scope = SCOPES.join(' ');

    const dcRes = await fetch(deviceUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            client_id: cid,
            scope
        })
    });
    const dc = await dcRes.json();
    if (!dcRes.ok) {
        throw new Error(
            `Device-Code fehlgeschlagen: ${dc.error || dcRes.status} ${dc.error_description || JSON.stringify(dc)}`
        );
    }

    console.log('');
    console.log('=== Microsoft-Anmeldung ===');
    console.log(dc.message || `Öffne ${dc.verification_uri} und gib Code ${dc.user_code} ein.`);
    console.log('');

    const intervalSec = Math.max(3, Number(dc.interval) || 5);
    const expiresAt = Date.now() + (Number(dc.expires_in) || 900) * 1000;

    while (Date.now() < expiresAt) {
        await sleep(intervalSec * 1000);
        const tr = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
                client_id: cid,
                device_code: dc.device_code
            })
        });
        const tok = await tr.json();
        if (tr.ok && tok.access_token) {
            return {
                accessToken: String(tok.access_token),
                refreshToken: tok.refresh_token ? String(tok.refresh_token) : '',
                expiresIn: Number(tok.expires_in) || 3600
            };
        }
        if (tok.error === 'authorization_pending') continue;
        if (tok.error === 'slow_down') {
            await sleep(intervalSec * 1000);
            continue;
        }
        throw new Error(
            `Token-Polling: ${tok.error || tr.status} ${tok.error_description || JSON.stringify(tok)}`
        );
    }
    throw new Error('Device-Code abgelaufen – bitte erneut starten.');
}

/**
 * @param {string} method
 * @param {string} pathOrUrl
 * @param {string} token
 * @param {any} [body]
 * @param {Record<string, string>} [extraHeaders]
 */
export async function graphJson(method, pathOrUrl, token, body, extraHeaders = {}) {
    const url = /^https?:\/\//i.test(pathOrUrl) ? pathOrUrl : GRAPH + pathOrUrl;
    /** @type {Record<string, string>} */
    const headers = {
        Authorization: `Bearer ${token}`,
        ...extraHeaders
    };
    let payload;
    if (body !== undefined && body !== null) {
        if (Buffer.isBuffer(body) || body instanceof Uint8Array || Readable.isReadable(body)) {
            payload = body;
        } else if (typeof body === 'string') {
            payload = body;
        } else {
            headers['Content-Type'] = headers['Content-Type'] || 'application/json';
            payload = JSON.stringify(body);
        }
    }
    const res = await fetch(url, { method, headers, body: payload });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = { raw: text };
        }
    }
    if (!res.ok) {
        const msg =
            (data && (data.error?.message || data.error_description)) ||
            text.slice(0, 400) ||
            res.statusText;
        const err = new Error(`Graph ${method} ${pathOrUrl} → ${res.status}: ${msg}`);
        /** @type {any} */ (err).status = res.status;
        /** @type {any} */ (err).data = data;
        throw err;
    }
    return data;
}

/**
 * @param {string} token
 * @param {string} groupId
 */
export async function getGroupDrive(token, groupId) {
    const drive = await graphJson('GET', `/groups/${encodeURIComponent(groupId)}/drive`, token);
    if (!drive?.id) throw new Error('Keine Drive für Gruppe ' + groupId);
    return drive;
}

/**
 * @param {string} token
 * @param {string} query
 */
export async function searchGroups(token, query) {
    const q = String(query || '').trim();
    if (!q) return [];
    if (/^[0-9a-f-]{36}$/i.test(q)) {
        const g = await graphJson(
            'GET',
            `/groups/${encodeURIComponent(q)}?$select=id,displayName,mail,mailNickname`,
            token
        );
        return g ? [g] : [];
    }
    const filter = `startswith(displayName,'${q.replace(/'/g, "''")}')`;
    const data = await graphJson(
        'GET',
        `/groups?$filter=${encodeURIComponent(filter)}&$select=id,displayName,mail,mailNickname&$top=25`,
        token
    );
    return (data && data.value) || [];
}

/**
 * @param {string} token
 * @param {string} teamName
 */
export async function resolveGroupIdByName(token, teamName) {
    const list = await searchGroups(token, teamName);
    if (!list.length) throw new Error(`Kein Team/Gruppe gefunden für „${teamName}“`);
    const exact = list.filter(
        (g) => String(g.displayName || '').toLowerCase() === teamName.toLowerCase()
    );
    const pick = exact[0] || list[0];
    if (list.length > 1 && exact.length !== 1) {
        console.warn(
            `Warnung: mehrere Treffer für „${teamName}“, nutze „${pick.displayName}“ (${pick.id}).`
        );
    }
    return pick.id;
}

/**
 * Pfad relativ zum Drive-Root anlegen (per Path-API, ohne Children-$top-Fallen).
 * @param {string} token
 * @param {string} driveId
 * @param {string} relativeFolderPath e.g. Import aus Nextcloud/Unterordner
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
            /** @type {any} */
            const err = e;
            if (err.status !== 404) throw e;
        }
        const parentId = last?.id || 'root';
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
            /** @type {any} */
            const err = e;
            if (err.status === 409) {
                last = await graphJson('GET', getUrl, token);
            } else {
                throw e;
            }
        }
    }
    return last;
}

const SMALL_UPLOAD_MAX = 4 * 1024 * 1024;

/**
 * @param {Readable} stream
 * @returns {Promise<Buffer>}
 */
async function streamToBuffer(stream) {
    /** @type {Buffer[]} */
    const chunks = [];
    for await (const chunk of stream) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    return Buffer.concat(chunks);
}

/**
 * @param {string} token
 * @param {string} driveId
 * @param {string} parentFolderId
 * @param {string} fileName
 * @param {Readable} stream
 * @param {number|null} size
 */
export async function uploadFile(token, driveId, parentFolderId, fileName, stream, size) {
    const name = String(fileName || '').trim();
    if (!name) throw new Error('Dateiname leer');

    const knownSize = size != null && size > 0 ? size : null;
    if (knownSize != null && knownSize < SMALL_UPLOAD_MAX) {
        const buf = await streamToBuffer(stream);
        return graphJson(
            'PUT',
            `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentFolderId)}:/${encodeURIComponent(name)}:/content`,
            token,
            buf,
            { 'Content-Type': 'application/octet-stream' }
        );
    }

    // Unbekannte oder große Dateien: Upload-Session (Chunked)
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

    const chunkSize = 5 * 1024 * 1024; // 5 MiB (Vielfaches von 320 KiB)
    let offset = 0;
    /** @type {Buffer[]} */
    let pending = [];
    let pendingLen = 0;
    let lastResult = null;

    /**
     * @param {Buffer} buf
     * @param {boolean} isLast
     */
    async function sendChunk(buf, isLast) {
        const start = offset;
        const end = offset + buf.length - 1;
        const total = knownSize != null ? String(knownSize) : isLast ? String(end + 1) : '*';
        const res = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Content-Length': String(buf.length),
                'Content-Range': `bytes ${start}-${end}/${total}`
            },
            body: buf
        });
        const text = await res.text();
        let data = null;
        try {
            data = text ? JSON.parse(text) : null;
        } catch {
            data = { raw: text };
        }
        if (!res.ok && res.status !== 202) {
            throw new Error(`Upload-Chunk ${res.status}: ${text.slice(0, 300)}`);
        }
        offset += buf.length;
        lastResult = data;
        return data;
    }

    for await (const chunk of stream) {
        const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
        pending.push(buf);
        pendingLen += buf.length;
        while (pendingLen >= chunkSize) {
            let need = chunkSize;
            /** @type {Buffer[]} */
            const parts = [];
            while (need > 0 && pending.length) {
                const head = pending[0];
                if (head.length <= need) {
                    parts.push(head);
                    need -= head.length;
                    pending.shift();
                } else {
                    parts.push(head.subarray(0, need));
                    pending[0] = head.subarray(need);
                    need = 0;
                }
            }
            pendingLen -= chunkSize;
            await sendChunk(Buffer.concat(parts), false);
        }
    }

    if (pendingLen > 0) {
        await sendChunk(Buffer.concat(pending), true);
    } else if (offset === 0) {
        // leere Datei
        await sendChunk(Buffer.alloc(0), true);
    }

    return lastResult;
}

/**
 * Retry-Hülle.
 * @template T
 * @param {() => Promise<T>} fn
 * @param {{ retries?: number, label?: string }} [opts]
 */
export async function withRetry(fn, opts = {}) {
    const retries = opts.retries ?? 4;
    let lastErr;
    for (let i = 0; i <= retries; i++) {
        try {
            return await fn();
        } catch (e) {
            lastErr = e;
            /** @type {any} */
            const err = e;
            const status = err.status || 0;
            const retryable = status === 429 || status >= 500 || status === 0;
            if (!retryable || i === retries) break;
            const wait = Math.min(60_000, 1000 * Math.pow(2, i));
            console.warn(
                `Retry ${i + 1}/${retries}${opts.label ? ' (' + opts.label + ')' : ''} in ${wait}ms: ${err.message || err}`
            );
            await sleep(wait);
        }
    }
    throw lastErr;
}
