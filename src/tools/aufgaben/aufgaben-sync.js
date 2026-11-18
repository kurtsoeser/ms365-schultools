/**
 * Live-Sync Aufgaben ↔ SharePoint (local-first).
 * 403/Fehler wipe’t lokalen Cache und Demo nicht.
 */

import {
    createAufgabeItem,
    deleteAufgabeItem,
    loadAllAufgaben,
    resolveAufgabenContext,
    updateAufgabeItem
} from './aufgaben-graph.js';
import { isDemoSeedItem } from './aufgaben-model.js';
import { resolveAufgabenSiteUrl } from './aufgaben-settings.js';
import {
    loadState,
    saveState,
    seedDemoIfEmpty,
    upsertAufgabe
} from './aufgaben-store.js';

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

export function getAufgabenSyncMeta() {
    return { ...syncMeta };
}

export function isAufgabenAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|accessDenied|Authorization_RequestDenied/i.test(msg);
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getAufgabenSyncMeta();
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function ensureAufgabenContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = resolveAufgabenSiteUrl();
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
    cachedCtx = await resolveAufgabenContext(siteUrl);
    setMeta({
        siteUrl: cachedCtx.webUrl,
        listTitle: (cachedCtx.list && cachedCtx.list.title) || '',
        connected: true
    });
    return cachedCtx;
}

/**
 * Pull von SharePoint. Bei Fehler: lokaler Stand bleibt.
 * @param {{ forceDemoIfEmpty?: boolean }} [opts]
 */
export async function pullAufgabenFromSharePoint(opts = {}) {
    const siteUrl = resolveAufgabenSiteUrl();
    if (!siteUrl) {
        const state = loadState();
        setMeta({
            ok: !!state.localDemoOnly,
            connected: false,
            source: state.localDemoOnly ? 'demo' : state.aufgaben.length ? 'local' : 'none',
            error: 'Keine Site-URL hinterlegt.',
            siteUrl: ''
        });
        return { ok: false, state, meta: getAufgabenSyncMeta() };
    }

    try {
        const { ctx, items } = await loadAllAufgaben({ webUrl: siteUrl });
        cachedCtx = ctx;
        const live = (items || []).filter((a) => a && (a.title || a.id));
        const next = saveState({
            aufgaben: live,
            localDemoOnly: false,
            updatedAt: new Date().toISOString()
        });
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            siteUrl: ctx.webUrl,
            listTitle: (ctx.list && ctx.list.title) || '',
            lastPullAt: new Date().toISOString()
        });
        return { ok: true, state: next, meta: getAufgabenSyncMeta() };
    } catch (e) {
        const denied = isAufgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || 'SharePoint nicht erreichbar.');
        let state = loadState();
        if (opts.forceDemoIfEmpty === true && !(state.aufgaben || []).length) {
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
            meta: getAufgabenSyncMeta(),
            softError: denied
                ? 'Kein SharePoint-Zugriff – lokal weiterarbeiten.'
                : msg
        };
    }
}

/**
 * Lokale Aufgabe nach SPO pushen (Create oder Update).
 * @param {object} aufgabe
 */
export async function pushAufgabe(aufgabe) {
    if (!aufgabe || isDemoSeedItem(aufgabe)) {
        return { ok: true, skipped: true, item: aufgabe };
    }
    try {
        const ctx = await ensureAufgabenContext();
        if (!ctx) {
            return {
                ok: false,
                item: aufgabe,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (aufgabe.spoItemId) {
            remote = await updateAufgabeItem(ctx, aufgabe);
        } else {
            remote = await createAufgabeItem(ctx, aufgabe);
        }
        const merged = {
            ...aufgabe,
            ...remote,
            id: aufgabe.id || remote.id,
            spoItemId: remote.spoItemId || aufgabe.spoItemId
        };
        const res = upsertAufgabe(merged);
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
        return { ok: true, item: res.item || merged, state: res.state };
    } catch (e) {
        const denied = isAufgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: aufgabe,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * @param {object} aufgabe
 */
export async function pushAufgabeDelete(aufgabe) {
    if (!aufgabe || isDemoSeedItem(aufgabe) || !aufgabe.spoItemId) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureAufgabenContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        await deleteAufgabeItem(ctx, aufgabe.spoItemId);
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isAufgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}
