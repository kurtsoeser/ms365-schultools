/** SessionStorage-Übergabe Import-Seite → Stammdaten / Einrichtung */

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

export function consumeWebuntisImportPayload() {
    try {
        const raw = sessionStorage.getItem(WEBUNTIS_IMPORT_PAYLOAD_KEY);
        if (!raw) return null;
        sessionStorage.removeItem(WEBUNTIS_IMPORT_PAYLOAD_KEY);
        return JSON.parse(raw);
    } catch {
        return null;
    }
}

export function resolveReturnUrl(fromParam) {
    const from = String(fromParam || '').toLowerCase();
    if (from === 'einrichtung') return '../einrichtung.html';
    if (from === 'tenant' || from === 'stammdaten') return '../tenant.html';
    try {
        const stored = sessionStorage.getItem(WEBUNTIS_IMPORT_RETURN_KEY);
        if (stored) return stored;
    } catch {
        /* ignore */
    }
    return '../tenant.html';
}
