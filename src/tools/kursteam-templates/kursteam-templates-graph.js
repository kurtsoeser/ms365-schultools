/**
 * Graph: Teams suchen, Kanäle lesen/anlegen/umbenennen (delegiert).
 */

export const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.Read.All',
    'https://graph.microsoft.com/Channel.ReadBasic.All',
    'https://graph.microsoft.com/Channel.Create',
    'https://graph.microsoft.com/ChannelSettings.ReadWrite.All',
    'https://graph.microsoft.com/Files.ReadWrite.All'
];

function G() {
    const api = window.ms365GraphUnifiedGroups;
    if (!api) throw new Error('Graph-Modul nicht geladen.');
    return api;
}

/**
 * @returns {Promise<string>}
 */
export async function getToken() {
    if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
        return window.ms365AuthAcquireTokenPopup(GRAPH_SCOPES);
    }
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return window.ms365AuthAcquireToken(GRAPH_SCOPES);
    }
    // Fallback: Unified-Groups-Token (ohne Channel-Scopes – Create kann scheitern)
    return G().getGraphToken();
}

/**
 * @param {string} method
 * @param {string} path
 * @param {string} token
 * @param {object} [body]
 */
export async function graphJson(method, path, token, body) {
    return G().graphJson(method, path, token, body);
}

function guidLike(s) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
        String(s || '').trim()
    );
}

/**
 * @param {string} query
 * @returns {Promise<Array<{ id: string, displayName: string, mail: string, mailNickname: string }>>}
 */
export async function searchTeams(query) {
    const q = String(query || '').trim();
    if (!q) throw new Error('Suchbegriff fehlt.');
    const token = await getToken();
    const api = G();

    if (guidLike(q)) {
        const g = await api.fetchGroup(token, q);
        if (!g || !g.id) throw new Error('Gruppe nicht gefunden.');
        return [
            {
                id: g.id,
                displayName: g.displayName || '',
                mail: g.mail || '',
                mailNickname: g.mailNickname || ''
            }
        ];
    }

    const list = await api.searchUnifiedGroups(token, q);
    return (list || []).map(function (g) {
        return {
            id: g.id,
            displayName: g.displayName || '',
            mail: g.mail || '',
            mailNickname: g.mailNickname || ''
        };
    });
}

/**
 * @param {string} teamId Group-/Team-ID
 * @returns {Promise<Array<{ id: string, displayName: string, membershipType: string }>>}
 */
export async function listChannels(teamId) {
    const id = String(teamId || '').trim();
    if (!id) throw new Error('Team-ID fehlt.');
    const token = await getToken();
    const out = [];
    let next =
        '/teams/' +
        encodeURIComponent(id) +
        '/channels?$select=id,displayName,membershipType';

    while (next) {
        const data = await graphJson('GET', next, token);
        const rows = (data && data.value) || [];
        for (const c of rows) {
            if (!c || !c.id) continue;
            out.push({
                id: String(c.id),
                displayName: String(c.displayName || ''),
                membershipType: String(c.membershipType || 'standard')
            });
        }
        next = (data && data['@odata.nextLink']) || null;
    }
    return out;
}

/**
 * @param {string} teamId
 * @param {string} displayName
 * @param {string} [description]
 */
export async function createChannel(teamId, displayName, description) {
    const id = String(teamId || '').trim();
    const name = String(displayName || '').trim();
    if (!id) throw new Error('Team-ID fehlt.');
    if (!name) throw new Error('Kanalname fehlt.');
    const token = await getToken();
    const body = {
        displayName: name,
        description: String(description || '').trim(),
        membershipType: 'standard'
    };
    return graphJson('POST', '/teams/' + encodeURIComponent(id) + '/channels', token, body);
}

/**
 * @param {string} teamId
 * @param {string} channelId
 * @param {string} displayName
 */
export async function renameChannel(teamId, channelId, displayName) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    const name = String(displayName || '').trim();
    if (!tid || !cid) throw new Error('Team- oder Kanal-ID fehlt.');
    if (!name) throw new Error('Neuer Kanalname fehlt.');
    const token = await getToken();
    return graphJson(
        'PATCH',
        '/teams/' + encodeURIComponent(tid) + '/channels/' + encodeURIComponent(cid),
        token,
        { displayName: name }
    );
}

/**
 * @param {import('./kursteam-templates-logic.js').DiffRow[]} rows
 * @param {string} teamId
 * @param {{
 *   doRename?: boolean,
 *   onProgress?: (info: { step: number, total: number, message: string, action: string, name: string }) => void
 * }} [opts]
 */
export async function applyDiff(rows, teamId, opts) {
    const doRename = !opts || opts.doRename !== false;
    const onProgress = opts && typeof opts.onProgress === 'function' ? opts.onProgress : null;
    const list = Array.isArray(rows) ? rows : [];
    /** @type {{ status: string, action: string, name: string, ok: boolean, error?: string }[]} */
    const results = [];

    const work = [];
    for (const row of list) {
        if (!row) continue;
        if (row.status === 'create' && row.templateChannel) work.push(row);
        else if (row.status === 'rename' && doRename && row.templateChannel && row.teamChannel) {
            work.push(row);
        }
    }
    const total = work.length;
    let step = 0;

    for (const row of work) {
        step += 1;
        if (row.status === 'create' && row.templateChannel) {
            const name = row.templateChannel.displayName;
            try {
                if (onProgress) {
                    onProgress({
                        step,
                        total,
                        action: 'create',
                        name,
                        message: 'Anlegen: ' + name
                    });
                }
                await createChannel(teamId, name);
                results.push({ status: 'create', action: 'create', name, ok: true });
            } catch (e) {
                const msg = (e && e.message) || String(e);
                // Conflict / bereits vorhanden
                if (/conflict|already exists|NameAlreadyExists/i.test(msg)) {
                    results.push({
                        status: 'create',
                        action: 'skip_exists',
                        name,
                        ok: true,
                        error: msg
                    });
                } else {
                    results.push({ status: 'create', action: 'create', name, ok: false, error: msg });
                }
            }
            continue;
        }
        if (row.status === 'rename' && row.templateChannel && row.teamChannel) {
            const name = row.templateChannel.displayName;
            try {
                if (onProgress) {
                    onProgress({
                        step,
                        total,
                        action: 'rename',
                        name,
                        message: 'Umbenennen: ' + row.teamChannel.displayName + ' → ' + name
                    });
                }
                await renameChannel(teamId, row.teamChannel.id, name);
                results.push({ status: 'rename', action: 'rename', name, ok: true });
            } catch (e) {
                results.push({
                    status: 'rename',
                    action: 'rename',
                    name,
                    ok: false,
                    error: (e && e.message) || String(e)
                });
            }
        }
    }
    return results;
}

/**
 * Dateien-Ordner eines Kanals (Drive-Ordner im Team).
 * @param {string} teamId
 * @param {string} channelId
 * @returns {Promise<{ driveId: string, folderId: string, webUrl: string }>}
 */
export async function getChannelFilesFolder(teamId, channelId) {
    const tid = String(teamId || '').trim();
    const cid = String(channelId || '').trim();
    if (!tid || !cid) throw new Error('Team- oder Kanal-ID fehlt.');
    const token = await getToken();
    const data = await graphJson(
        'GET',
        '/teams/' +
            encodeURIComponent(tid) +
            '/channels/' +
            encodeURIComponent(cid) +
            '/filesFolder?$select=id,name,webUrl,parentReference',
        token
    );
    const driveId = data && data.parentReference && data.parentReference.driveId;
    const folderId = data && data.id;
    if (!driveId || !folderId) {
        throw new Error('Dateien-Ordner des Kanals nicht gefunden (Team hat ggf. noch kein SharePoint).');
    }
    return {
        driveId: String(driveId),
        folderId: String(folderId),
        webUrl: (data && data.webUrl) || ''
    };
}

/**
 * Kleine Datei (unter 4 MB Graph-Simple-Upload) in den Kanal-Dateien-Ordner legen.
 * @param {string} teamId
 * @param {string} channelId
 * @param {{ name: string, bytes: Uint8Array|ArrayBuffer, contentType?: string }} file
 */
export async function uploadFileToChannel(teamId, channelId, file) {
    const name = String((file && file.name) || '').trim();
    if (!name) throw new Error('Dateiname fehlt.');
    const bytes =
        file && file.bytes instanceof ArrayBuffer
            ? new Uint8Array(file.bytes)
            : file && file.bytes
              ? file.bytes
              : null;
    if (!bytes || !bytes.byteLength) throw new Error('Dateiinhalt fehlt.');
    if (bytes.byteLength > 4 * 1024 * 1024) {
        throw new Error('Datei ist größer als 4 MB (Test-Upload).');
    }
    const folder = await getChannelFilesFolder(teamId, channelId);
    const token = await getToken();
    const path =
        '/drives/' +
        encodeURIComponent(folder.driveId) +
        '/items/' +
        encodeURIComponent(folder.folderId) +
        ':/' +
        encodeURIComponent(name) +
        ':/content';
    const url = 'https://graph.microsoft.com/v1.0' + path;
    const res = await fetch(url, {
        method: 'PUT',
        headers: {
            Authorization: 'Bearer ' + token,
            'Content-Type': (file && file.contentType) || 'application/octet-stream'
        },
        body: bytes
    });
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
            data && data.error && data.error.message
                ? data.error.message
                : text || 'HTTP ' + res.status;
        throw new Error('Upload: ' + msg);
    }
    return {
        id: data && data.id ? String(data.id) : '',
        name: data && data.name ? String(data.name) : name,
        webUrl: (data && data.webUrl) || folder.webUrl || '',
        size: Number(data && data.size) || bytes.byteLength
    };
}

