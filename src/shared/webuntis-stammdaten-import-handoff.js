/** SessionStorage-Übergabe Import-Seite → Schulregister */

export const WEBUNTIS_IMPORT_PAYLOAD_KEY = 'ms365WebuntisImportPayload';
export const WEBUNTIS_IMPORT_RETURN_KEY = 'ms365WebuntisImportReturn';

export function stashWebuntisImportPayload(payload, returnTo) {
    try {
        sessionStorage.setItem(WEBUNTIS_IMPORT_PAYLOAD_KEY, JSON.stringify(payload || {}));
        if (returnTo) sessionStorage.setItem(WEBUNTIS_IMPORT_RETURN_KEY, String(returnTo));
    } catch {
        /* ignore */
    }
}

export function peekWebuntisImportPayload() {
    try {
        const raw = sessionStorage.getItem(WEBUNTIS_IMPORT_PAYLOAD_KEY);
        if (!raw) return null;
        return JSON.parse(raw);
    } catch {
        return null;
    }
}

export function clearWebuntisImportPayload() {
    try {
        sessionStorage.removeItem(WEBUNTIS_IMPORT_PAYLOAD_KEY);
        sessionStorage.removeItem(WEBUNTIS_IMPORT_RETURN_KEY);
    } catch {
        /* ignore */
    }
}

export function consumeWebuntisImportPayload() {
    const payload = peekWebuntisImportPayload();
    if (payload) clearWebuntisImportPayload();
    return payload;
}

export function resolveReturnUrl(fromParam) {
    const from = String(fromParam || '').toLowerCase();
    if (from === 'playbook-daten-import') return 'playbook-daten-import-verknuepfen.html';
    if (from === 'tenant' || from === 'stammdaten' || from === 'einrichtung') return '../tenant.html';
    try {
        const stored = sessionStorage.getItem(WEBUNTIS_IMPORT_RETURN_KEY);
        if (stored) return stored;
    } catch {
        /* ignore */
    }
    return '../tenant.html';
}
