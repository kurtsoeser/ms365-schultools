/**
 * Lokaler Demo-/Pilot-Store (localStorage) – gleiche Daten für Admin + Formular
 * im selben Browser. Produktion: Azure Functions Gateway.
 */

import {
    STATUS,
    applyGeprueftStamp,
    applyImportedStamp,
    attemptUnlock,
    createAccessPin,
    createEnrollmentRecord,
    hashPin,
    isParentEditLocked,
    normalizePayload,
    normalizeStatus,
    publicRecordView
} from './stammdaten-portal-logic.js';
import { resolveFieldProfile } from './stammdaten-portal-fields.js';

const STORAGE_KEY = 'ms365-schueler-stammdaten-portal-v1';
const SESSION_KEY = 'ms365-schueler-stammdaten-sessions-v1';

function loadRoot() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return defaultRoot();
        const parsed = JSON.parse(raw);
        return {
            ...defaultRoot(),
            ...parsed,
            records: Array.isArray(parsed.records) ? parsed.records : []
        };
    } catch {
        return defaultRoot();
    }
}

/** Pilot-Defaults für zentrales Gateway (ohne Admin-Key – der bleibt lokal). */
export const PILOT_REMOTE_DEFAULTS = Object.freeze({
    schoolName: 'MS365 Schule',
    schoolSlug: 'ms365schule',
    apiBaseUrl: 'https://func-ms365-ssp-dev.azurewebsites.net/api',
    spSiteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
    spListName: 'SchuelerAnmeldungen',
    mode: 'remote',
    presetId: 'kern'
});

function defaultRoot() {
    return {
        schoolName: PILOT_REMOTE_DEFAULTS.schoolName,
        schoolSlug: PILOT_REMOTE_DEFAULTS.schoolSlug,
        formBaseUrl: '',
        apiBaseUrl: PILOT_REMOTE_DEFAULTS.apiBaseUrl,
        functionKey: '',
        spSiteUrl: PILOT_REMOTE_DEFAULTS.spSiteUrl,
        spListName: PILOT_REMOTE_DEFAULTS.spListName,
        mode: PILOT_REMOTE_DEFAULTS.mode,
        presetId: PILOT_REMOTE_DEFAULTS.presetId,
        fieldOverrides: {},
        pepper: 'local-demo-pepper',
        records: []
    };
}

/**
 * Leere/veraltete Local-Pilot-Werte auf Remote-Gateway umbiegen (ohne Key zu überschreiben).
 * @param {object} [patch]
 */
export function ensurePilotRemoteDefaults(patch = {}) {
    const root = loadRoot();
    const next = { ...root };
    let changed = false;
    for (const [k, v] of Object.entries(PILOT_REMOTE_DEFAULTS)) {
        const cur = String(next[k] ?? '').trim();
        const shouldFill =
            !cur ||
            (k === 'mode' && cur === 'local') ||
            (k === 'schoolSlug' && cur === 'demo') ||
            (k === 'schoolName' && (cur === 'Demo-Schule' || cur === 'MS365-Schul-Tools Demo'));
        if (shouldFill && String(v) !== cur) {
            next[k] = v;
            changed = true;
        }
    }
    if (patch && typeof patch === 'object') {
        for (const [k, v] of Object.entries(patch)) {
            if (v != null && String(v).trim() !== '' && String(next[k] || '') !== String(v)) {
                next[k] = v;
                changed = true;
            }
        }
    }
    if (changed) saveRoot(next);
    return loadPortalSettings();
}

function saveRoot(root) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(root));
    return root;
}

export function loadPortalSettings() {
    const root = loadRoot();
    return {
        schoolName: root.schoolName,
        schoolSlug: root.schoolSlug,
        formBaseUrl: root.formBaseUrl,
        apiBaseUrl: root.apiBaseUrl,
        functionKey: root.functionKey || '',
        spSiteUrl: root.spSiteUrl || '',
        spListName: root.spListName || 'SchuelerAnmeldungen',
        mode: root.mode || 'local',
        presetId: root.presetId || 'kern',
        fieldOverrides: root.fieldOverrides || {},
        pepper: root.pepper
    };
}

export function savePortalSettings(patch) {
    const root = loadRoot();
    Object.assign(root, patch || {});
    return saveRoot(root);
}

export function getFieldProfileFromSettings(settings = loadPortalSettings()) {
    return resolveFieldProfile(settings.presetId || 'kern', settings.fieldOverrides || {});
}

export function listRecords() {
    return loadRoot().records.map((r) => ({
        uniqueGuid: r.uniqueGuid,
        title: r.title,
        status: normalizeStatus(r.status),
        createdAt: r.createdAt,
        updatedAt: r.updatedAt,
        reviewedAt: r.reviewedAt || null,
        reviewedBy: r.reviewedBy || '',
        reviewedByName: r.reviewedByName || '',
        payload: normalizePayload(r.payload)
    }));
}

export function getRecordByGuid(guid) {
    return loadRoot().records.find((r) => r.uniqueGuid === String(guid || '')) || null;
}

export async function createLocalRecord({ title, payload } = {}) {
    const root = loadRoot();
    const { record, accessPin } = await createEnrollmentRecord({
        title,
        payload,
        pepper: root.pepper
    });
    root.records.unshift(record);
    saveRoot(root);
    return { record: publicRecordView(record), accessPin };
}

/**
 * Admin: neuen PIN setzen (alter PIN ist nicht wiederherstellbar, nur Hash).
 * Entsperrt gleichzeitig fehlgeschlagene Versuche.
 * @param {string} guid
 */
export async function resetLocalAccessPin(guid) {
    const root = loadRoot();
    const idx = root.records.findIndex((r) => r.uniqueGuid === String(guid || ''));
    if (idx < 0) throw new Error('Datensatz nicht gefunden.');
    const accessPin = createAccessPin();
    const accessPinHash = await hashPin(accessPin, root.pepper);
    root.records[idx] = {
        ...root.records[idx],
        accessPinHash,
        failedAttempts: 0,
        lockedUntil: null,
        updatedAt: new Date().toISOString()
    };
    saveRoot(root);
    return {
        accessPin,
        record: publicRecordView(root.records[idx]),
        formUrl: buildParentFormUrl(root.records[idx].uniqueGuid, root)
    };
}

/** Admin-Ansicht inkl. Payload (ohne PIN-Hash). */
export function getAdminRecordDetail(guid) {
    const record = getRecordByGuid(guid);
    if (!record) return null;
    const view = publicRecordView(record);
    return {
        ...view,
        reviewedAt: record.reviewedAt || view.reviewedAt || null,
        reviewedBy: record.reviewedBy || view.reviewedBy || '',
        reviewedByName: record.reviewedByName || view.reviewedByName || '',
        formUrl: buildParentFormUrl(record.uniqueGuid)
    };
}

export async function localUnlock(guid, pin) {
    const root = loadRoot();
    const idx = root.records.findIndex((r) => r.uniqueGuid === String(guid || ''));
    if (idx < 0) {
        return {
            ok: false,
            message:
                'Ungültiger Zugang (lokal kein Datensatz). Wenn der Eintrag über das Gateway/SharePoint angelegt wurde: Formular mit mode=remote öffnen.'
        };
    }
    const result = await attemptUnlock(root.records[idx], pin, { pepper: root.pepper });
    if (result.record) {
        root.records[idx] = result.record;
        saveRoot(root);
    }
    if (!result.ok) return { ok: false, message: result.message || 'Ungültiger Zugang.' };
    const token = await issueSession(guid);
    return {
        ok: true,
        token,
        record: publicRecordView(root.records[idx]),
        profile: getFieldProfileFromSettings(root)
    };
}

export function localLoad(token) {
    const session = readSession(token);
    if (!session) return { ok: false, message: 'Sitzung abgelaufen. Bitte erneut anmelden.' };
    const record = getRecordByGuid(session.guid);
    if (!record) return { ok: false, message: 'Ungültiger Zugang.' };
    return {
        ok: true,
        record: publicRecordView(record),
        profile: getFieldProfileFromSettings()
    };
}

export function localSave(token, payload, { submit = false } = {}) {
    const session = readSession(token);
    if (!session) return { ok: false, message: 'Sitzung abgelaufen. Bitte erneut anmelden.' };
    const root = loadRoot();
    const idx = root.records.findIndex((r) => r.uniqueGuid === session.guid);
    if (idx < 0) return { ok: false, message: 'Ungültiger Zugang.' };
    const current = root.records[idx];
    if (isParentEditLocked(current.status)) {
        return {
            ok: false,
            code: 'locked_after_review',
            message:
                'Dieser Datensatz wurde vom Sekretariat geprüft und kann nicht mehr geändert werden. Bei Korrekturen bitte die Schule kontaktieren.'
        };
    }
    const nextPayload = normalizePayload(payload);
    root.records[idx] = {
        ...current,
        payload: nextPayload,
        title:
            [nextPayload.schueler.familienname, nextPayload.schueler.vorname].filter(Boolean).join(', ') ||
            current.title,
        status: submit ? STATUS.VOLLSTAENDIG : STATUS.IN_BEARBEITUNG,
        updatedAt: new Date().toISOString()
    };
    saveRoot(root);
    return { ok: true, record: publicRecordView(root.records[idx]) };
}

/**
 * @param {string[]} guids
 * @param {{ actorEmail?: string, actorName?: string }} [actor]
 */
export function markImported(guids, actor = {}) {
    void actor;
    const set = new Set((guids || []).map(String));
    const root = loadRoot();
    const touched = [];
    root.records = root.records.map((r) => {
        if (!set.has(r.uniqueGuid)) return r;
        const result = applyImportedStamp(r);
        if (!result.ok) return r;
        touched.push(result.record);
        return result.record;
    });
    saveRoot(root);
    return touched;
}

/**
 * Sekretariats-Stempel: freigegeben für Sokrates-Export.
 * @param {string[]} guids
 * @param {{ actorEmail?: string, actorName?: string }} [actor]
 */
export function markGeprueft(guids, actor = {}) {
    const set = new Set((guids || []).map(String));
    const root = loadRoot();
    const touched = [];
    root.records = root.records.map((r) => {
        if (!set.has(r.uniqueGuid)) return r;
        const result = applyGeprueftStamp(r, actor);
        if (!result.ok) return r;
        touched.push(result.record);
        return result.record;
    });
    saveRoot(root);
    return touched;
}

/** Roh-Root für Demo-Seed (nur lokal). */
export function __loadRootForDemo() {
    return loadRoot();
}

/** @param {object} root */
export function __saveRootForDemo(root) {
    return saveRoot(root);
}

async function issueSession(guid) {
    const tokenBytes = new Uint8Array(24);
    crypto.getRandomValues(tokenBytes);
    const token = [...tokenBytes].map((b) => b.toString(16).padStart(2, '0')).join('');
    const sessions = loadSessions();
    sessions[token] = { guid: String(guid), exp: Date.now() + 60 * 60 * 1000 };
    localStorage.setItem(SESSION_KEY, JSON.stringify(sessions));
    return token;
}

function loadSessions() {
    try {
        return JSON.parse(localStorage.getItem(SESSION_KEY) || '{}') || {};
    } catch {
        return {};
    }
}

function readSession(token) {
    const sessions = loadSessions();
    const s = sessions[String(token || '')];
    if (!s) return null;
    if (Date.now() > Number(s.exp || 0)) {
        delete sessions[token];
        localStorage.setItem(SESSION_KEY, JSON.stringify(sessions));
        return null;
    }
    return s;
}

export function buildParentFormUrl(guid, settings = loadPortalSettings()) {
    const base =
        String(settings.formBaseUrl || '').trim() ||
        new URL('../tools/schueler-stammdaten-formular.html', window.location.href).href;
    const url = new URL(base, window.location.href);
    url.searchParams.set('id', guid);
    if (settings.schoolSlug) url.searchParams.set('schule', settings.schoolSlug);
    const mode = settings.mode === 'remote' ? 'remote' : 'local';
    url.searchParams.set('mode', mode);
    // Damit Eltern-Formular remote den Gateway findet (auch ohne Tenant-JSON).
    if (mode === 'remote' && settings.apiBaseUrl) {
        url.searchParams.set('api', String(settings.apiBaseUrl).replace(/\/$/, ''));
    }
    return url.toString();
}

export function exportSettingsAsPublicConfig(settings = loadPortalSettings()) {
    return {
        schoolName: settings.schoolName,
        schoolSlug: settings.schoolSlug,
        mode: settings.mode === 'remote' ? 'remote' : 'local',
        apiBaseUrl: settings.apiBaseUrl || '',
        presetId: settings.presetId || 'kern',
        fieldOverrides: settings.fieldOverrides || {},
        privacyUrl: settings.privacyUrl || ''
        // functionKey bewusst nicht öffentlich
    };
}
