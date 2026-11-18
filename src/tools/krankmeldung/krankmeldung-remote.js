/**
 * Optional Remote-Gateway für Krankmeldungen (gleiche Basis-URL wie Einsammler-API möglich).
 */

import { loadKrankmeldungSettings } from './krankmeldung-settings.js';

function base() {
    return loadKrankmeldungSettings().apiBaseUrl;
}

export function krankmeldungRemoteConfigured() {
    return !!base();
}

/**
 * @param {string} path
 * @param {{ method?: string, body?: unknown, admin?: boolean }} [opts]
 */
async function apiFetch(path, opts = {}) {
    const root = base();
    if (!root) throw new Error('Keine Krankmeldung-API konfiguriert.');
    const headers = { Accept: 'application/json', 'Content-Type': 'application/json' };
    if (opts.admin) {
        const key = loadKrankmeldungSettings().adminKey;
        if (key) headers['x-admin-key'] = key;
    }
    const res = await fetch(root + path, {
        method: opts.method || 'GET',
        headers,
        body: opts.body != null ? JSON.stringify(opts.body) : undefined
    });
    const text = await res.text();
    let data = null;
    try {
        data = text ? JSON.parse(text) : null;
    } catch {
        data = { message: text };
    }
    if (!res.ok) {
        throw new Error((data && (data.error || data.message)) || 'HTTP ' + res.status);
    }
    return data || {};
}

/**
 * @param {object} meldung
 */
export async function remoteSubmitMeldung(meldung) {
    return apiFetch('/krankmeldung', { method: 'POST', body: { meldung } });
}

export async function remoteListMeldungen() {
    return apiFetch('/krankmeldung/list', { method: 'GET', admin: true });
}

export async function remoteHealth() {
    return apiFetch('/krankmeldung/health');
}
