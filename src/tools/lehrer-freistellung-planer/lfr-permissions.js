/**
 * Entra-Gruppen, Einzelpersonen & Direktion-Genehmiger für Lehrer-Freistellungen.
 */
import {
    normalizePlannerUsers,
    accountIsPlannerUserInList
} from '../freistellung-planer/freistellung-planer-direktion-users.js';

const STORAGE_KEY = 'ms365-lfr-perms-v1';
const SETUP_KEY = 'ms365-lfr-setup-v1';

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
        /** @type {Array<{ id: string, displayName: string, mail: string }>} */
        lehrerUsers: normalizePlannerUsers(x.lehrerUsers),
        /** @type {Array<{ id: string, displayName: string, mail: string }>} */
        direktionUsers: normalizePlannerUsers(x.direktionUsers),
        /** Power-Automate-Genehmiger (eine E-Mail) – getrennt vom Listen-/Planer-Zugriff */
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

/** Lehrer-Freistellungen-Setup-Cfg (Schritt 2) */
export function direktionEmailFromLfrSetup() {
    try {
        const setup = JSON.parse(localStorage.getItem(SETUP_KEY) || '{}');
        return String(setup.emailDirektion || '').trim().toLowerCase();
    } catch {
        return '';
    }
}

/**
 * Effektive Genehmiger-Mail: Perms → LFR-Setup → Schüler-Setup.
 */
export function resolveLfrApproverEmail(perms) {
    const p = perms || loadLfrPermissions();
    return (
        String(p.emailDirektion || '').trim().toLowerCase() ||
        direktionEmailFromLfrSetup() ||
        direktionEmailFromFreistellungSetup()
    );
}

/**
 * Schreibt Genehmiger-Mail in Perms und Setup-Cfg (gemeinsame Quelle für Flow + UI).
 * @param {string} email
 */
export function persistLfrApproverEmail(email) {
    const mail = String(email || '')
        .trim()
        .toLowerCase();
    const perms = saveLfrPermissions({ ...loadLfrPermissions(), emailDirektion: mail });
    try {
        const setup = JSON.parse(localStorage.getItem(SETUP_KEY) || '{}') || {};
        setup.emailDirektion = mail;
        localStorage.setItem(SETUP_KEY, JSON.stringify(setup));
    } catch {
        /* ignore */
    }
    return perms;
}

/**
 * @param {string} accountEmail
 * @param {ReturnType<typeof normalizeLfrPermissions>} [perms]
 */
export function accountIsLfrDirektionAccess(accountEmail, perms) {
    const p = perms || loadLfrPermissions();
    const em = String(accountEmail || '')
        .trim()
        .toLowerCase();
    if (!em) return false;
    const approver = resolveLfrApproverEmail(p);
    if (approver && approver === em) return true;
    return accountIsPlannerUserInList(em, p.direktionUsers);
}

/**
 * @param {string} accountEmail
 * @param {ReturnType<typeof normalizeLfrPermissions>} [perms]
 */
export function accountIsLfrLehrerAccess(accountEmail, perms) {
    const p = perms || loadLfrPermissions();
    return accountIsPlannerUserInList(accountEmail, p.lehrerUsers);
}

export function lfrPermsStepConfigured(perms) {
    const p = normalizeLfrPermissions(perms || loadLfrPermissions());
    return !!(
        p.groupLehrerId ||
        p.groupDirektionId ||
        (p.lehrerUsers && p.lehrerUsers.length) ||
        (p.direktionUsers && p.direktionUsers.length)
    );
}
