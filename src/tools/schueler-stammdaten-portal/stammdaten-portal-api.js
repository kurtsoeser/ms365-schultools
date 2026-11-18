/**
 * Client für Eltern-API: local Store oder Remote-Gateway.
 */

import { localLoad, localSave, localUnlock } from './stammdaten-portal-store.js';

const SSP_AZURE_HOST_RE = /func-ms365-ssp-dev\.azurewebsites\.net/i;

function resolveApiBase(config = {}) {
    const configured = String(config.apiBaseUrl || '').replace(/\/$/, '');
    if (typeof window === 'undefined') return configured;
    const host = window.location.hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1';
    // Vite-Proxy /__ssp-gw → kein Browser-CORS
    if (isLocal && (!configured || SSP_AZURE_HOST_RE.test(configured))) {
        return `${window.location.origin}/__ssp-gw/api`;
    }
    return configured;
}

/**
 * @param {{ mode?: string, apiBaseUrl?: string }} config
 */
export function createPortalApi(config = {}) {
    const mode = config.mode === 'remote' ? 'remote' : 'local';
    const base = resolveApiBase(config);

    return {
        mode,
        async unlock(id, pin) {
            if (mode === 'local') return localUnlock(id, pin);
            if (!base) {
                return { ok: false, message: 'API-Base-URL fehlt (Modus remote). Bitte Einrichtung prüfen.' };
            }
            return postJson(`${base}/unlock`, { id, pin });
        },
        async load(token) {
            if (mode === 'local') return localLoad(token);
            return getJson(`${base}/load`, token);
        },
        async save(token, payload, { submit = false } = {}) {
            if (mode === 'local') return localSave(token, payload, { submit });
            return postJson(`${base}/save`, { payload, submit }, token);
        }
    };
}

async function postJson(url, body, token) {
    let res;
    try {
        res = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                ...(token ? { Authorization: `Bearer ${token}` } : {})
            },
            body: JSON.stringify(body || {})
        });
    } catch {
        return {
            ok: false,
            message:
                'Server vorübergehend nicht erreichbar. Bitte später erneut versuchen – Ihre eingegebenen Daten bleiben in diesem Formular.'
        };
    }
    const data = await safeJson(res);
    if (!res.ok) {
        return {
            ok: false,
            code: data?.code || '',
            message: softParentMessage(res.status, data)
        };
    }
    return { ok: true, ...data };
}

async function getJson(url, token) {
    let res;
    try {
        res = await fetch(url, {
            method: 'GET',
            headers: {
                ...(token ? { Authorization: `Bearer ${token}` } : {})
            }
        });
    } catch {
        return {
            ok: false,
            message: 'Server vorübergehend nicht erreichbar. Bitte später erneut versuchen.'
        };
    }
    const data = await safeJson(res);
    if (!res.ok) {
        return { ok: false, message: softParentMessage(res.status, data) };
    }
    return { ok: true, ...data };
}

function softParentMessage(status, data) {
    const detail = String(data?.error || data?.message || '').trim();
    if (status === 409 || data?.code === 'locked_after_review') {
        return (
            detail ||
            'Dieser Datensatz wurde vom Sekretariat geprüft und kann nicht mehr geändert werden.'
        );
    }
    if (status === 401) {
        return detail || 'Sitzung abgelaufen oder PIN ungültig. Bitte erneut öffnen.';
    }
    if (status === 403) {
        return detail || 'Kein Zugriff. Bitte den Link der Schule verwenden.';
    }
    if (status >= 500) {
        return detail || 'Server vorübergehend nicht erreichbar. Bitte später erneut versuchen.';
    }
    return detail || 'Anfrage fehlgeschlagen.';
}

async function safeJson(res) {
    try {
        return await res.json();
    } catch {
        return {};
    }
}
