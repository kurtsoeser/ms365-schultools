/**
 * Client für election-api (mandantenweite anonyme Stimmen).
 */

function apiConfig() {
    return (typeof window !== 'undefined' && window.MS365_ELECTION_API) || {};
}

export function electionApiConfigured() {
    return !!String(apiConfig().baseUrl || '').trim();
}

function baseUrl() {
    return String(apiConfig().baseUrl || '').replace(/\/+$/, '');
}

function authHeaders(accessToken) {
    const headers = {
        Accept: 'application/json',
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + String(accessToken || '').trim()
    };
    const key = String(apiConfig().functionKey || '').trim();
    if (key) headers['x-functions-key'] = key;
    return headers;
}

async function acquireElectionToken(opts) {
    const primary = String(apiConfig().scope || '').trim();
    const fallback = String(apiConfig().fallbackScope || '').trim();
    const scopes = [primary, fallback].filter(Boolean);
    const wantPopup = !!(opts && opts.popup);
    let lastErr = null;

    for (const scope of scopes) {
        try {
            if (wantPopup && typeof window.ms365AuthAcquireTokenPopup === 'function') {
                return await window.ms365AuthAcquireTokenPopup([scope]);
            }
            if (typeof window.ms365AuthAcquireToken === 'function') {
                return await window.ms365AuthAcquireToken([scope]);
            }
            if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
                return await window.ms365AuthAcquireTokenPopup([scope]);
            }
        } catch (e) {
            lastErr = e;
        }
    }
    throw lastErr || new Error('Bitte zuerst mit MS365 anmelden (Election-API-Scope).');
}

async function parseResponse(res) {
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = { message: text };
        }
    }
    if (!res.ok) {
        const msg = (data && (data.error || data.message)) || 'HTTP ' + res.status;
        const err = new Error(String(msg));
        err.status = res.status;
        err.payload = data;
        err.code = data && data.code ? String(data.code) : '';
        throw err;
    }
    return data || {};
}

/**
 * @param {string} path
 * @param {{ method?: string, body?: unknown, popup?: boolean }} [opts]
 */
async function apiFetch(path, opts) {
    const base = baseUrl();
    if (!base) throw new Error('MS365_ELECTION_API.baseUrl ist nicht gesetzt.');
    const token = await acquireElectionToken({ popup: !!(opts && opts.popup) });
    const res = await fetch(base + path, {
        method: (opts && opts.method) || 'GET',
        headers: authHeaders(token),
        body: opts && opts.body !== undefined ? JSON.stringify(opts.body) : undefined
    });
    return parseResponse(res);
}

export async function apiListElections() {
    return apiFetch('', { popup: true });
}

export async function apiGetElection(electionId) {
    return apiFetch('/' + encodeURIComponent(electionId));
}

export async function apiUpsertElection(election) {
    return apiFetch('/' + encodeURIComponent(election.id), {
        method: 'PUT',
        body: { election },
        popup: true
    });
}

export async function apiDeleteElection(electionId) {
    return apiFetch('/' + encodeURIComponent(electionId), { method: 'DELETE', popup: true });
}

export async function apiElectionMe(electionId) {
    return apiFetch('/' + encodeURIComponent(electionId) + '/me');
}

export async function apiCastVote(electionId, ballot) {
    return apiFetch('/' + encodeURIComponent(electionId) + '/vote', {
        method: 'POST',
        body: ballot,
        popup: true
    });
}

export async function apiListBallots(electionId, roundId) {
    return apiFetch(
        '/' + encodeURIComponent(electionId) + '/ballots?roundId=' + encodeURIComponent(roundId),
        { popup: true }
    );
}

export { acquireElectionToken };
