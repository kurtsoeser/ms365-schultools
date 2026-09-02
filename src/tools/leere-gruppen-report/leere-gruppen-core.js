export function compareDe(a, b) {
    return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
}

export function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

export function csvEscape(cell) {
    const s = String(cell ?? '');
    if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
}

export function rowsToCsv(rows, columns) {
    const header = columns.map((c) => c.label).join(';');
    const lines = rows.map((row) => columns.map((c) => csvEscape(c.value(row))).join(';'));
    return '\uFEFF' + header + '\n' + lines.join('\n');
}

export function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
}

export async function getGraphToken(scopes) {
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return await window.ms365AuthAcquireToken(scopes);
    }
    throw new Error('Bitte oben rechts anmelden (MSAL-Widget nicht verfügbar).');
}

export async function graphRequest(method, pathOrUrl, token, body, extraHeaders) {
    const url = pathOrUrl.indexOf('http') === 0 ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
    let attempt = 0;
    while (true) {
        try {
            const headers = { Authorization: 'Bearer ' + token };
            if (extraHeaders && typeof extraHeaders === 'object') Object.assign(headers, extraHeaders);
            let payload = undefined;
            if (body !== undefined) {
                headers['Content-Type'] = 'application/json';
                payload = JSON.stringify(body);
            }
            const res = await fetch(url, { method, headers, body: payload });
            if ((res.status === 429 || res.status === 503) && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || String(Math.min(60, 2 ** attempt + 2)), 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        } catch (e) {
            if (attempt < 8) {
                await sleep(Math.min(30, 2 ** attempt + 1) * 1000);
                attempt++;
                continue;
            }
            throw e;
        }
    }
}

export async function graphJson(method, pathOrUrl, token, body, extraHeaders) {
    const res = await graphRequest(method, pathOrUrl, token, body, extraHeaders);
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error
                ? JSON.stringify(data.error)
                : text || String(res.status);
        const err = new Error(method + ' ' + pathOrUrl + ': ' + msg);
        err.status = res.status;
        throw err;
    }
    return data || {};
}

export async function fetchAllPages(token, initialPath, onProgress, extraHeaders) {
    const out = [];
    let next = initialPath;
    let page = 0;
    while (next) {
        page++;
        const data = await graphJson('GET', next, token, undefined, extraHeaders);
        const vals = data.value;
        if (Array.isArray(vals)) for (let i = 0; i < vals.length; i++) out.push(vals[i]);
        next = data['@odata.nextLink'] || null;
        if (typeof onProgress === 'function') onProgress({ page, loaded: out.length, hasMore: !!next });
    }
    return out;
}

export function parseCountValue(body) {
    if (typeof body === 'number' && Number.isFinite(body)) return Math.max(0, Math.trunc(body));
    if (typeof body === 'boolean') return body ? 1 : 0;
    if (body == null) return -1;
    const n = parseInt(String(body).trim(), 10);
    return isNaN(n) ? -1 : Math.max(0, n);
}

export async function fetchCount(token, groupId, segment) {
    try {
        const path = '/groups/' + encodeURIComponent(groupId) + '/' + segment + '/$count';
        const res = await graphRequest('GET', path, token, undefined, { ConsistencyLevel: 'eventual' });
        const text = await res.text();
        if (!res.ok) return -1;
        return parseCountValue(text);
    } catch {
        return -1;
    }
}

export async function runPool(tasks, concurrency) {
    const results = new Array(tasks.length);
    let i = 0;
    async function worker() {
        while (true) {
            const idx = i++;
            if (idx >= tasks.length) return;
            results[idx] = await tasks[idx]();
        }
    }
    const n = Math.max(1, Math.min(concurrency, tasks.length || 1));
    const workers = [];
    for (let w = 0; w < n; w++) workers.push(worker());
    await Promise.all(workers);
    return results;
}

async function postGraphBatch(token, requests) {
    const data = await graphJson('POST', '/$batch', token, { requests });
    return Array.isArray(data.responses) ? data.responses : [];
}

/**
 * Zählt Besitzer/Mitglieder für viele Gruppen über Graph $batch (max. 20 Requests/Batch).
 * Einzelne Fehlschläge liefern -1 und brechen die Gesamtanalyse nicht ab.
 *
 * @param {string} token
 * @param {Array<object>} groups
 * @param {(p: { done: number, total: number }) => void} [onProgress]
 * @param {{ getToken?: () => Promise<string>, concurrency?: number }} [opts]
 */
export async function enrichGroupsWithCounts(token, groups, onProgress, opts) {
    const list = Array.isArray(groups) ? groups : [];
    const total = list.length;
    if (!total) return [];

    const getToken = opts && typeof opts.getToken === 'function' ? opts.getToken : null;
    const concurrency = Math.max(1, Math.min((opts && opts.concurrency) || 2, 4));
    const groupsPerBatch = 10; // 2 Requests je Gruppe → max. 20 / Batch

    const rows = new Array(total);
    let done = 0;
    let currentToken = token;

    async function refreshToken() {
        if (!getToken) return currentToken;
        try {
            currentToken = await getToken();
        } catch {
            /* bestehendes Token weiterverwenden */
        }
        return currentToken;
    }

    function reportProgress() {
        if (typeof onProgress === 'function') onProgress({ done, total });
    }

    const chunks = [];
    for (let i = 0; i < total; i += groupsPerBatch) {
        chunks.push({ start: i, items: list.slice(i, i + groupsPerBatch) });
    }

    async function countsForChunkFallback(items) {
        const map = new Map();
        for (const g of items) {
            const id = String(g && g.id ? g.id : '');
            if (!id) continue;
            const [owners, members] = await Promise.all([
                fetchCount(currentToken, id, 'owners'),
                fetchCount(currentToken, id, 'members')
            ]);
            map.set(id, { owners, members });
        }
        return map;
    }

    async function countsForChunk(items) {
        const map = new Map();
        for (const g of items) {
            const id = String(g && g.id ? g.id : '');
            if (id) map.set(id, { owners: -1, members: -1 });
        }

        /** @type {Array<{ id: string, groupId: string, kind: 'owners' | 'members' }>} */
        const meta = [];
        const requests = [];
        for (let i = 0; i < items.length; i++) {
            const id = String(items[i] && items[i].id ? items[i].id : '');
            if (!id) continue;
            // Batch-Request-IDs: max. 36 Zeichen (GUID allein schon 36) → kurze Indizes
            const oid = 'o' + i;
            const mid = 'm' + i;
            meta.push({ id: oid, groupId: id, kind: 'owners' }, { id: mid, groupId: id, kind: 'members' });
            requests.push(
                {
                    id: oid,
                    method: 'GET',
                    url: '/groups/' + encodeURIComponent(id) + '/owners/$count',
                    headers: { ConsistencyLevel: 'eventual' }
                },
                {
                    id: mid,
                    method: 'GET',
                    url: '/groups/' + encodeURIComponent(id) + '/members/$count',
                    headers: { ConsistencyLevel: 'eventual' }
                }
            );
        }
        if (!requests.length) return map;

        const byReqId = new Map(meta.map((m) => [m.id, m]));

        await refreshToken();
        let responses;
        try {
            responses = await postGraphBatch(currentToken, requests);
        } catch {
            return await countsForChunkFallback(items);
        }

        const retry = [];
        for (let r = 0; r < responses.length; r++) {
            const resp = responses[r] || {};
            const info = byReqId.get(String(resp.id || ''));
            if (!info) continue;
            const entry = map.get(info.groupId);
            if (!entry) continue;
            const status = Number(resp.status) || 0;
            if (status === 429 || status === 503) {
                retry.push(info);
                continue;
            }
            if (status >= 200 && status < 300) entry[info.kind] = parseCountValue(resp.body);
        }

        for (let t = 0; t < retry.length; t++) {
            const item = retry[t];
            const entry = map.get(item.groupId);
            if (!entry) continue;
            entry[item.kind] = await fetchCount(currentToken, item.groupId, item.kind);
        }
        return map;
    }

    await runPool(
        chunks.map((chunk) => async () => {
            let map;
            try {
                map = await countsForChunk(chunk.items);
            } catch {
                map = new Map();
                for (const g of chunk.items) {
                    const id = String(g && g.id ? g.id : '');
                    if (id) map.set(id, { owners: -1, members: -1 });
                }
            }
            for (let j = 0; j < chunk.items.length; j++) {
                const g = chunk.items[j];
                const id = String(g && g.id ? g.id : '');
                if (!id) {
                    rows[chunk.start + j] = null;
                } else {
                    const c = map.get(id) || { owners: -1, members: -1 };
                    rows[chunk.start + j] = buildRow(g, c.owners, c.members);
                }
                done++;
            }
            reportProgress();
        }),
        concurrency
    );

    return rows.filter(Boolean);
}

export function isUnified(g) {
    const gt = g && Array.isArray(g.groupTypes) ? g.groupTypes : [];
    return gt.indexOf('Unified') !== -1;
}

export function isTeam(g) {
    const ro = g && Array.isArray(g.resourceProvisioningOptions) ? g.resourceProvisioningOptions : [];
    return ro.indexOf('Team') !== -1;
}

export function isSecurity(g) {
    return !!(g && g.securityEnabled && !g.mailEnabled && !isUnified(g));
}

export function isMailGroup(g) {
    return !!(g && g.mailEnabled && !isUnified(g));
}

export function groupKindLabel(g) {
    const parts = [];
    if (isUnified(g)) parts.push('M365');
    if (isTeam(g)) parts.push('Team');
    if (g && g.securityEnabled && !g.mailEnabled) parts.push('Sicherheit');
    if (g && g.mailEnabled && !isUnified(g)) parts.push('Mail');
    return parts.length ? parts.join(' · ') : 'Sonstige';
}

export function kindBadgesHtml(row) {
    const out = [];
    if (row.isUnified) out.push('<span class="lgr-badge is-m365">M365</span>');
    if (row.isTeam) out.push('<span class="lgr-badge is-team">Team</span>');
    if (row.isSecurity) out.push('<span class="lgr-badge is-security">Sicherheit</span>');
    if (row.isMail) out.push('<span class="lgr-badge is-mail">Mail</span>');
    if (!out.length) out.push('<span class="lgr-badge">Sonstige</span>');
    return out.join('');
}

/**
 * Konvertiert Graph-Group-Objekte (mit zusätzlich gezählten owners/members) zu Tabellenzeilen.
 */
export function buildRow(g, owners, members) {
    const flags = [];
    if (owners === 0) flags.push('ohne Besitzer');
    if (members === 0) flags.push('ohne Mitglieder');
    if (owners < 0 || members < 0) flags.push('Zählen fehlgeschlagen');
    return {
        id: String(g.id || ''),
        displayName: String(g.displayName || ''),
        mail: String(g.mail || ''),
        kind: groupKindLabel(g),
        isUnified: isUnified(g),
        isTeam: isTeam(g),
        isSecurity: isSecurity(g),
        isMail: isMailGroup(g),
        owners,
        members,
        flags: flags.join(', ') || '–'
    };
}

export function buildGroupsListInitialPath(scopeMode) {
    const select =
        'id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions,visibility,securityEnabled,mailEnabled';
    // $top=100: stabile Pagination (bei großen Tenants und Advanced Queries zuverlässiger als 999)
    if (scopeMode === 'all') {
        return {
            path: '/groups?$select=' + encodeURIComponent(select) + '&$top=100',
            headers: {}
        };
    }
    if (scopeMode === 'team') {
        return {
            path:
                '/groups?$filter=' +
                encodeURIComponent("resourceProvisioningOptions/Any(x:x eq 'Team')") +
                '&$select=' +
                encodeURIComponent(select) +
                '&$count=true&$top=100',
            headers: { ConsistencyLevel: 'eventual' }
        };
    }
    return {
        path:
            '/groups?$filter=' +
            encodeURIComponent("groupTypes/any(c:c eq 'Unified')") +
            '&$select=' +
            encodeURIComponent(select) +
            '&$count=true&$top=100',
        headers: { ConsistencyLevel: 'eventual' }
    };
}
