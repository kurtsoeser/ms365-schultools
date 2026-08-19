'use strict';

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

async function graphRequest(method, path, token, body) {
    const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
    let attempt = 0;
    while (true) {
        const headers = { Authorization: 'Bearer ' + token };
        if (body !== undefined) {
            headers['Content-Type'] = 'application/json';
        }
        const res = await fetch(url, {
            method,
            headers,
            body: body !== undefined ? JSON.stringify(body) : undefined
        });
        if (res.status === 429 && attempt < 8) {
            const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
            await sleep((Number.isNaN(ra) ? 5 : ra) * 1000);
            attempt++;
            continue;
        }
        return res;
    }
}

async function graphJson(method, path, token, body) {
    const res = await graphRequest(method, path, token, body);
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
        throw new Error(method + ' ' + path + ': ' + msg);
    }
    return data || {};
}

function isGraphDuplicateRefError(err) {
    const msg = String(err && err.message ? err.message : err);
    return /already exist/i.test(msg) || /already exists/i.test(msg);
}

function parseTeamsOperationPath(locationHeader) {
    if (!locationHeader) return null;
    const loc = String(locationHeader).trim();
    const m = loc.match(/teams\('([^']+)'\)\/operations\('([^']+)'\)/i);
    if (m) return '/teams/' + m[1] + '/operations/' + m[2];
    const m2 = loc.match(/\/teams\/([^/]+)\/operations\/([^/?\s]+)/i);
    if (m2) return '/teams/' + m2[1] + '/operations/' + m2[2];
    return null;
}

async function pollTeamsAsyncOperation(token, operationPath, log) {
    const maxAttempts = 120;
    for (let i = 0; i < maxAttempts; i++) {
        await sleep(2000);
        const data = await graphJson('GET', operationPath, token, undefined);
        const st = String(data.status || data.Status || '').toLowerCase();
        if (st === 'succeeded') {
            log('Teams: Bereitstellung abgeschlossen (Template educationClass).');
            return;
        }
        if (st === 'failed') {
            const errMsg =
                (data.error && (data.error.message || JSON.stringify(data.error))) ||
                JSON.stringify(data);
            throw new Error('Team-Bereitstellung fehlgeschlagen: ' + errMsg);
        }
        if (i > 0 && i % 8 === 0) {
            log('Teams: Warte auf Bereitstellung … (' + i * 2 + ' s)');
        }
    }
    throw new Error('Timeout: Team-Bereitstellung (async) nicht abgeschlossen.');
}

module.exports = {
    sleep,
    graphRequest,
    graphJson,
    isGraphDuplicateRefError,
    parseTeamsOperationPath,
    pollTeamsAsyncOperation
};
