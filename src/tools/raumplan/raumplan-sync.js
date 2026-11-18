/**
 * Live-Sync Raumplan ↔ SharePoint (local-first).
 * Plan-Bild/PDF bleiben lokal. 403/Fehler wipe’t Cache und Demo nicht.
 */

import {
    createRaumItem,
    deleteRaumItem,
    loadAllRaeume,
    resolveRaumplanContext,
    updateRaumItem
} from './raumplan-graph.js';
import { isDemoSeedItem } from './raumplan-model.js';
import { resolveRaumplanSiteUrl } from './raumplan-settings.js';
import {
    loadState,
    saveState,
    seedDemoIfEmpty,
    upsertRaum
} from './raumplan-store.js';

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

export function getRaumplanSyncMeta() {
    return { ...syncMeta };
}

export function isRaumplanAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|accessDenied|Authorization_RequestDenied/i.test(msg);
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getRaumplanSyncMeta();
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function ensureRaumplanContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = resolveRaumplanSiteUrl();
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
    cachedCtx = await resolveRaumplanContext(siteUrl);
    setMeta({
        siteUrl: cachedCtx.webUrl,
        listTitle: (cachedCtx.list && cachedCtx.list.title) || '',
        connected: true
    });
    return cachedCtx;
}

/**
 * Pull von SharePoint. Plan-Assets und lokaler Stand bei Fehler bleiben.
 * @param {{ forceDemoIfEmpty?: boolean }} [opts]
 */
export async function pullRaumplanFromSharePoint(opts = {}) {
    const siteUrl = resolveRaumplanSiteUrl();
    if (!siteUrl) {
        let state = loadState();
        if (opts.forceDemoIfEmpty !== false && !(state.raeume || []).length) {
            state = seedDemoIfEmpty();
        }
        setMeta({
            ok: !!state.localDemoOnly,
            connected: false,
            source: state.localDemoOnly ? 'demo' : state.raeume.length ? 'local' : 'none',
            error: 'Keine Site-URL hinterlegt.',
            siteUrl: ''
        });
        return { ok: false, state, meta: getRaumplanSyncMeta() };
    }

    try {
        const { ctx, items } = await loadAllRaeume({ webUrl: siteUrl });
        cachedCtx = ctx;
        const live = (items || []).filter((r) => r && (r.code || r.name || r.id));
        const prev = loadState();
        const next = saveState({
            raeume: live,
            planImageDataUrl: prev.planImageDataUrl,
            planPdfName: prev.planPdfName,
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
        return { ok: true, state: next, meta: getRaumplanSyncMeta() };
    } catch (e) {
        const denied = isRaumplanAccessDeniedError(e);
        const msg = String((e && e.message) || e || 'SharePoint nicht erreichbar.');
        let state = loadState();
        if (opts.forceDemoIfEmpty !== false && !(state.raeume || []).length) {
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
            meta: getRaumplanSyncMeta(),
            softError: denied
                ? 'Kein SharePoint-Zugriff – lokal weiterarbeiten.'
                : msg
        };
    }
}

/**
 * Lokalen Raum nach SPO pushen (Create oder Update). Demo-Seeds werden übersprungen.
 * @param {object} raum
 */
export async function pushRaumplanRaum(raum) {
    if (!raum || isDemoSeedItem(raum)) {
        return { ok: true, skipped: true, item: raum };
    }
    try {
        const ctx = await ensureRaumplanContext();
        if (!ctx) {
            return {
                ok: false,
                item: raum,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (raum.spoItemId) {
            remote = await updateRaumItem(ctx, raum);
        } else {
            remote = await createRaumItem(ctx, raum);
        }
        const merged = {
            ...raum,
            ...remote,
            id: raum.id || remote.id,
            spoItemId: remote.spoItemId || raum.spoItemId
        };
        const res = upsertRaum(merged);
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
        const denied = isRaumplanAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: raum,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * @param {object} raum
 */
export async function pushRaumplanDelete(raum) {
    if (!raum || isDemoSeedItem(raum) || !raum.spoItemId) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureRaumplanContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        await deleteRaumItem(ctx, raum.spoItemId);
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isRaumplanAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}
