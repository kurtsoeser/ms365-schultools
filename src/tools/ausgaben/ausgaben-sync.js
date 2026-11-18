/**
 * Live-Sync Ausgabenverwaltung ↔ SharePoint (local-first).
 * Inventar + Vorgänge. 403/Fehler wipe’t lokalen Cache und Demo nicht.
 * Typen bleiben lokal (UI); werden beim Pull aus Inventar ergänzt.
 */

import {
    createGegenstandItem,
    createVorgangItem,
    deleteGegenstandItem,
    deleteVorgangItem,
    loadAllAusgaben,
    resolveAusgabenContext,
    updateGegenstandItem,
    updateVorgangItem
} from './ausgaben-graph.js';
import { isDemoSeedItem, normalizeTyp } from './ausgaben-model.js';
import { resolveAusgabenSiteUrl } from './ausgaben-settings.js';
import { loadState, saveState, seedDemoIfEmpty } from './ausgaben-store.js';

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

export function getAusgabenSyncMeta() {
    return { ...syncMeta };
}

export function isAusgabenAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|accessDenied|Authorization_RequestDenied/i.test(msg);
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getAusgabenSyncMeta();
}

/**
 * @param {object[]} gegenstaende
 * @param {object[]} existingTypen
 */
export function mergeTypenFromGegenstaende(gegenstaende, existingTypen) {
    const byId = new Map();
    (existingTypen || []).forEach((t) => {
        const n = normalizeTyp(t);
        if (n.id) byId.set(n.id, n);
    });
    (gegenstaende || []).forEach((g) => {
        const id = String((g && g.typ) || '')
            .trim()
            .toLowerCase();
        if (!id || byId.has(id)) return;
        byId.set(
            id,
            normalizeTyp({
                id,
                label: String((g && g.typLabel) || id).trim() || id,
                icon: 'bi-tag'
            })
        );
    });
    return Array.from(byId.values());
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function ensureAusgabenContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = resolveAusgabenSiteUrl();
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
    cachedCtx = await resolveAusgabenContext(siteUrl);
    const titles = [
        cachedCtx.lists && cachedCtx.lists.gegenstaende && cachedCtx.lists.gegenstaende.title,
        cachedCtx.lists && cachedCtx.lists.vorgaenge && cachedCtx.lists.vorgaenge.title
    ]
        .filter(Boolean)
        .join(' + ');
    setMeta({
        siteUrl: cachedCtx.webUrl,
        listTitle: titles,
        connected: true
    });
    return cachedCtx;
}

/**
 * Pull von SharePoint. Bei Fehler: lokaler Stand bleibt.
 * @param {{ forceDemoIfEmpty?: boolean }} [opts]
 */
export async function pullAusgabenFromSharePoint(opts = {}) {
    const siteUrl = resolveAusgabenSiteUrl();
    if (!siteUrl) {
        const state = seedDemoIfEmpty();
        setMeta({
            ok: !!state.localDemoOnly,
            connected: false,
            source: state.localDemoOnly ? 'demo' : 'local',
            error: 'Keine Site-URL hinterlegt.',
            siteUrl: ''
        });
        return { ok: false, state, meta: getAusgabenSyncMeta() };
    }

    try {
        const { ctx, gegenstaende, vorgaenge } = await loadAllAusgaben({ webUrl: siteUrl });
        cachedCtx = ctx;
        const liveGg = (gegenstaende || []).filter((g) => g && (g.kennung || g.id));
        const liveVg = (vorgaenge || []).filter((v) => v && (v.id || v.gegenstandId));
        const prev = loadState();
        const next = saveState({
            typen: mergeTypenFromGegenstaende(liveGg, prev.typen || []),
            gegenstaende: liveGg,
            vorgaenge: liveVg,
            localDemoOnly: false,
            updatedAt: new Date().toISOString()
        });
        const titles = [
            ctx.lists && ctx.lists.gegenstaende && ctx.lists.gegenstaende.title,
            ctx.lists && ctx.lists.vorgaenge && ctx.lists.vorgaenge.title
        ]
            .filter(Boolean)
            .join(' + ');
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            siteUrl: ctx.webUrl,
            listTitle: titles,
            lastPullAt: new Date().toISOString()
        });
        return { ok: true, state: next, meta: getAusgabenSyncMeta() };
    } catch (e) {
        const denied = isAusgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || 'SharePoint nicht erreichbar.');
        let state = loadState();
        if (
            opts.forceDemoIfEmpty !== false &&
            !(state.gegenstaende || []).length &&
            !(state.typen || []).length
        ) {
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
            meta: getAusgabenSyncMeta(),
            softError: denied
                ? 'Kein SharePoint-Zugriff – lokal weiterarbeiten.'
                : msg
        };
    }
}

function markLiveAfterPush(ctx) {
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
}

/**
 * Lokalen Gegenstand nach SPO pushen (Create oder Update).
 * @param {object} gegenstand
 */
export async function pushAusgabenGegenstand(gegenstand) {
    if (!gegenstand || isDemoSeedItem(gegenstand)) {
        return { ok: true, skipped: true, item: gegenstand };
    }
    try {
        const ctx = await ensureAusgabenContext();
        if (!ctx) {
            return {
                ok: false,
                item: gegenstand,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (gegenstand.spoItemId) {
            remote = await updateGegenstandItem(ctx, gegenstand);
        } else {
            remote = await createGegenstandItem(ctx, gegenstand);
        }
        const merged = {
            ...gegenstand,
            ...remote,
            id: gegenstand.id || remote.id,
            spoItemId: remote.spoItemId || gegenstand.spoItemId
        };
        const state = loadState();
        const idx = state.gegenstaende.findIndex((x) => x.id === merged.id);
        if (idx >= 0) state.gegenstaende[idx] = { ...state.gegenstaende[idx], ...merged };
        else state.gegenstaende.unshift(merged);
        saveState(state);
        markLiveAfterPush(ctx);
        return { ok: true, item: merged, state: loadState() };
    } catch (e) {
        const denied = isAusgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: gegenstand,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * @param {object} gegenstand
 */
export async function pushAusgabenGegenstandDelete(gegenstand) {
    if (!gegenstand || isDemoSeedItem(gegenstand) || !gegenstand.spoItemId) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureAusgabenContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        await deleteGegenstandItem(ctx, gegenstand.spoItemId);
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isAusgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}

/**
 * Lokalen Vorgang nach SPO pushen (Create oder Update).
 * @param {object} vorgang
 */
export async function pushAusgabenVorgang(vorgang) {
    if (!vorgang || isDemoSeedItem(vorgang)) {
        return { ok: true, skipped: true, item: vorgang };
    }
    try {
        const ctx = await ensureAusgabenContext();
        if (!ctx) {
            return {
                ok: false,
                item: vorgang,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (vorgang.spoItemId) {
            remote = await updateVorgangItem(ctx, vorgang);
        } else {
            remote = await createVorgangItem(ctx, vorgang);
        }
        const merged = {
            ...vorgang,
            ...remote,
            id: vorgang.id || remote.id,
            spoItemId: remote.spoItemId || vorgang.spoItemId
        };
        const state = loadState();
        const idx = state.vorgaenge.findIndex((x) => x.id === merged.id);
        if (idx >= 0) state.vorgaenge[idx] = { ...state.vorgaenge[idx], ...merged };
        else state.vorgaenge.unshift(merged);
        saveState(state);
        markLiveAfterPush(ctx);
        return { ok: true, item: merged, state: loadState() };
    } catch (e) {
        const denied = isAusgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: vorgang,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * @param {object} vorgang
 */
export async function pushAusgabenVorgangDelete(vorgang) {
    if (!vorgang || isDemoSeedItem(vorgang) || !vorgang.spoItemId) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureAusgabenContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        await deleteVorgangItem(ctx, vorgang.spoItemId);
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isAusgabenAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}
