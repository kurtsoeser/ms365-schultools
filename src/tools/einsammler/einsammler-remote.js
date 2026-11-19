/**
 * Optional: Remote-Gateway für Public-Antworten (wenn apiBaseUrl gesetzt).
 */

import { loadEinsammlerSettings, remoteConfigured } from './einsammler-settings.js';

function base() {
    return loadEinsammlerSettings().apiBaseUrl;
}

/** Remote nur mit apiBaseUrl + linkSecret (a5); Default/Demo: aus (w2). */
export function einsammlerRemoteConfigured() {
    return remoteConfigured();
}

/**
 * @param {string} path
 * @param {{ method?: string, body?: unknown, admin?: boolean }} [opts]
 */
async function apiFetch(path, opts = {}) {
    const root = base();
    if (!root) throw new Error('Keine Einsammler-API konfiguriert.');
    const headers = { Accept: 'application/json', 'Content-Type': 'application/json' };
    if (opts.admin) {
        const key = loadEinsammlerSettings().adminKey;
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
 * @param {object[]} packs public payloads
 */
export async function remotePublishPacks(packs) {
    return apiFetch('/einsammler/publish', { method: 'POST', admin: true, body: { packs } });
}

/**
 * @param {string} token encoded public token
 */
export async function remoteView(token) {
    return apiFetch('/einsammler/view?p=' + encodeURIComponent(token));
}

/**
 * @param {{ token: string, status: string, kommentar?: string, anhangName?: string, anhangMime?: string, anhangSize?: number, anhangDataUrl?: string }} body
 */
export async function remoteRespond(body) {
    return apiFetch('/einsammler/respond', { method: 'POST', body });
}

export async function remoteHealth() {
    return apiFetch('/einsammler/health');
}
