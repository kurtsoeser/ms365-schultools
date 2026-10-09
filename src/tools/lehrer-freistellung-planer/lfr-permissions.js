/**
 * Entra-Gruppen & Direktion für Lehrer-Freistellungen (lokal + optional SharePoint-JSON später).
 */
const STORAGE_KEY = 'ms365-lfr-perms-v1';

/**
 * @returns {object}
 */
export function loadLfrPermissions() {
    try {
        const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
        return normalizeLfrPermissions(raw);
    } catch {
        return normalizeLfrPermissions({});
    }
}

/**
 * @param {object} raw
 */
export function normalizeLfrPermissions(raw) {
    const x = raw && typeof raw === 'object' ? raw : {};
    return {
        groupLehrerId: String(x.groupLehrerId || '').trim(),
        groupLehrerName: String(x.groupLehrerName || '').trim(),
        groupDirektionId: String(x.groupDirektionId || '').trim(),
        groupDirektionName: String(x.groupDirektionName || '').trim(),
        emailDirektion: String(x.emailDirektion || '').trim().toLowerCase()
    };
}

/**
 * @param {object} cfg
 */
export function saveLfrPermissions(cfg) {
    const n = normalizeLfrPermissions(cfg);
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(n));
    } catch {
        /* ignore */
    }
    return n;
}

/** Freistellungen-Setup (Schüler) als Fallback für Direktions-Mail */
export function direktionEmailFromFreistellungSetup() {
    try {
        const setup = JSON.parse(localStorage.getItem('ms365-freistellung-setup-v1') || '{}');
        return String(setup.emailDirektion || '').trim().toLowerCase();
    } catch {
        return '';
    }
}
