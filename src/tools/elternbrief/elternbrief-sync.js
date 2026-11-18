/**
 * Live-Sync Elternbrief ↔ SharePoint (local-first).
 * 403/Fehler wipe’t lokalen Cache und Demo nicht.
 * Public-Magic-Link schreibt nicht direkt nach SPO (kein Eltern-Login).
 */

import {
    createBriefItem,
    createEmpfaengerItem,
    deleteBriefItem,
    deleteEmpfaengerItem,
    loadAllElternbrief,
    resolveElternbriefContext,
    updateBriefItem,
    updateEmpfaengerItem
} from './elternbrief-graph.js';
import { isDemoSeedItem } from './elternbrief-model.js';
import { resolveElternbriefSiteUrl } from './elternbrief-settings.js';
import {
    loadState,
    saveState,
    seedDemoIfEmpty,
    upsertBrief,
    upsertEmpfaenger
} from './elternbrief-store.js';

/** @type {{
 *   ok: boolean,
 *   source: 'live'|'local'|'demo'|'none',
 *   error: string,
 *   siteUrl: string,
 *   listTitle: string,
 *   lastPullAt: string,
 *   connected: boolean
 * }} */
let syncMeta = {
    ok: false,
    source: 'none',
    error: '',
    siteUrl: '',
    listTitle: '',
    lastPullAt: '',
    connected: false
};

/** @type {object|null} */
let cachedCtx = null;

export function getElternbriefSyncMeta() {
    return { ...syncMeta };
}

export function isElternbriefAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|accessDenied|Authorization_RequestDenied/i.test(msg);
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getElternbriefSyncMeta();
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function ensureElternbriefContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = resolveElternbriefSiteUrl();
    if (!siteUrl) {
        cachedCtx = null;
        setMeta({
            ok: false,
            connected: false,
            source: syncMeta.source === 'demo' ? 'demo' : 'none',
            error: 'Keine Site-URL – Listen-Assistent oder System nutzen.',
            siteUrl: ''
        });
        return null;
    }
    cachedCtx = await resolveElternbriefContext(siteUrl);
    setMeta({
        siteUrl: cachedCtx.webUrl,
        listTitle:
            (cachedCtx.listBriefe && cachedCtx.listBriefe.title) ||
            (cachedCtx.listLese && cachedCtx.listLese.title) ||
            '',
        connected: true
    });
    return cachedCtx;
}

/**
 * Pull von SharePoint. Bei Fehler: lokaler Stand bleibt.
 * @param {{ forceDemoIfEmpty?: boolean }} [opts]
 */
export async function pullElternbriefFromSharePoint(opts = {}) {
    const siteUrl = resolveElternbriefSiteUrl();
    if (!siteUrl) {
        const state = seedDemoIfEmpty();
        setMeta({
            ok: !!state.localDemoOnly,
            connected: false,
            source: state.localDemoOnly ? 'demo' : 'local',
            error: 'Keine Site-URL hinterlegt.',
            siteUrl: ''
        });
        return { ok: false, state, meta: getElternbriefSyncMeta() };
    }

    try {
        const { ctx, briefe, empfaenger } = await loadAllElternbrief({ webUrl: siteUrl });
        cachedCtx = ctx;
        const liveBriefe = (briefe || []).filter((b) => b && (b.betreff || b.id));
        const liveEmp = (empfaenger || []).filter((e) => e && e.briefId);
        const next = saveState({
            briefe: liveBriefe,
            empfaenger: liveEmp,
            localDemoOnly: false,
            updatedAt: new Date().toISOString()
        });
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            siteUrl: ctx.webUrl,
            listTitle: (ctx.listBriefe && ctx.listBriefe.title) || '',
            lastPullAt: new Date().toISOString()
        });
        return { ok: true, state: next, meta: getElternbriefSyncMeta() };
    } catch (e) {
        const denied = isElternbriefAccessDeniedError(e);
        const msg = String((e && e.message) || e || 'SharePoint nicht erreichbar.');
        let state = loadState();
        if (opts.forceDemoIfEmpty !== false && !(state.briefe || []).length) {
            state = seedDemoIfEmpty();
        }
        setMeta({
            ok: false,
            connected: false,
            source: state.localDemoOnly ? 'demo' : 'local',
            error: denied ? '' : msg,
            siteUrl,
            lastPullAt: ''
        });
        return {
            ok: false,
            state,
            meta: getElternbriefSyncMeta(),
            softError: denied
                ? 'Kein SharePoint-Zugriff – lokal weiterarbeiten.'
                : msg
        };
    }
}

/**
 * Lokalen Brief nach SPO pushen (Create oder Update).
 * @param {object} brief
 */
export async function pushElternbriefBrief(brief) {
    if (!brief || isDemoSeedItem(brief)) {
        return { ok: true, skipped: true, item: brief };
    }
    try {
        const ctx = await ensureElternbriefContext();
        if (!ctx) {
            return {
                ok: false,
                item: brief,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (brief.spoItemId) {
            remote = await updateBriefItem(ctx, brief);
        } else {
            remote = await createBriefItem(ctx, brief);
        }
        const merged = {
            ...brief,
            ...remote,
            id: brief.id || remote.id,
            spoItemId: remote.spoItemId || brief.spoItemId
        };
        const res = upsertBrief(merged);
        const state = loadState();
        if (state.localDemoOnly) {
            state.localDemoOnly = false;
            saveState(state);
        }
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            siteUrl: ctx.webUrl
        });
        return { ok: true, item: (res && res.item) || merged, state: res && res.state };
    } catch (e) {
        const denied = isElternbriefAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: brief,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * Empfänger / Lesebestätigung nach SPO pushen.
 * @param {object} emp
 */
export async function pushElternbriefEmpfaenger(emp) {
    if (!emp || isDemoSeedItem(emp)) {
        return { ok: true, skipped: true, item: emp };
    }
    try {
        const ctx = await ensureElternbriefContext();
        if (!ctx) {
            return {
                ok: false,
                item: emp,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (emp.spoItemId) {
            remote = await updateEmpfaengerItem(ctx, emp);
        } else {
            remote = await createEmpfaengerItem(ctx, emp);
        }
        const merged = {
            ...emp,
            ...remote,
            id: emp.id || remote.id,
            spoItemId: remote.spoItemId || emp.spoItemId
        };
        const res = upsertEmpfaenger(merged);
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            siteUrl: ctx.webUrl
        });
        return { ok: true, item: (res && res.item) || merged, state: res && res.state };
    } catch (e) {
        const denied = isElternbriefAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: emp,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * Alle Empfänger eines Briefs pushen (nach Speichern / Status).
 * @param {string} briefId
 */
export async function pushElternbriefEmpfaengerForBrief(briefId) {
    const bid = String(briefId || '').trim();
    const state = loadState();
    const rows = (state.empfaenger || []).filter((e) => e.briefId === bid);
    /** @type {{ ok: boolean, skipped?: boolean, error?: string }[]} */
    const results = [];
    for (const emp of rows) {
        results.push(await pushElternbriefEmpfaenger(emp));
    }
    const failed = results.find((r) => !r.ok && !r.skipped);
    return {
        ok: !failed,
        count: rows.length,
        error: failed ? failed.error : '',
        results
    };
}

/**
 * @param {object} brief
 * @param {object[]} [empfaenger]
 */
export async function pushElternbriefDelete(brief, empfaenger) {
    if (!brief || isDemoSeedItem(brief)) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureElternbriefContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        const emps = Array.isArray(empfaenger)
            ? empfaenger
            : (loadState().empfaenger || []).filter((e) => e.briefId === brief.id);
        for (const emp of emps) {
            if (emp && emp.spoItemId && !isDemoSeedItem(emp)) {
                await deleteEmpfaengerItem(ctx, emp.spoItemId);
            }
        }
        if (brief.spoItemId) {
            await deleteBriefItem(ctx, brief.spoItemId);
        }
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isElternbriefAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}
