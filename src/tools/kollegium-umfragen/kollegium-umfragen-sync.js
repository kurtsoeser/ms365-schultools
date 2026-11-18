/**
 * Live-Sync Kollegiumsumfragen ↔ SharePoint (local-first).
 * 403/Fehler wipe’t lokalen Cache und Demo nicht.
 */

import {
    createAntwortItem,
    createUmfrageItem,
    deleteAntwortItem,
    deleteUmfrageItem,
    loadAllKollegium,
    resolveKollegiumContext,
    updateAntwortItem,
    updateUmfrageItem
} from './kollegium-umfragen-graph.js';
import { isDemoSeedItem } from './kollegium-umfragen-model.js';
import { resolveKollegiumSiteUrl } from './kollegium-umfragen-settings.js';
import {
    loadState,
    saveState,
    seedDemoIfEmpty,
    upsertAntwort,
    upsertUmfrage
} from './kollegium-umfragen-store.js';

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

export function getKollegiumSyncMeta() {
    return { ...syncMeta };
}

export function isKollegiumAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|accessDenied|Authorization_RequestDenied/i.test(msg);
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getKollegiumSyncMeta();
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function ensureKollegiumContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = resolveKollegiumSiteUrl();
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
    cachedCtx = await resolveKollegiumContext(siteUrl);
    const uTitle = (cachedCtx.listUmfragen && cachedCtx.listUmfragen.title) || '';
    const aTitle = (cachedCtx.listAntworten && cachedCtx.listAntworten.title) || '';
    setMeta({
        siteUrl: cachedCtx.webUrl,
        listTitle: uTitle + (aTitle ? (uTitle ? ' + ' : '') + aTitle : ''),
        connected: true
    });
    return cachedCtx;
}

/**
 * Pull von SharePoint. Bei Fehler: lokaler Stand bleibt.
 * @param {{ forceDemoIfEmpty?: boolean }} [opts]
 */
export async function pullKollegiumFromSharePoint(opts = {}) {
    const siteUrl = resolveKollegiumSiteUrl();
    if (!siteUrl) {
        const state = seedDemoIfEmpty();
        setMeta({
            ok: !!state.localDemoOnly,
            connected: false,
            source: state.localDemoOnly ? 'demo' : 'local',
            error: 'Keine Site-URL hinterlegt.',
            siteUrl: ''
        });
        return { ok: false, state, meta: getKollegiumSyncMeta() };
    }

    try {
        const { ctx, umfragen, antworten } = await loadAllKollegium({ webUrl: siteUrl });
        cachedCtx = ctx;
        const liveUmfragen = (umfragen || []).filter((u) => u && (u.titel || u.id));
        const liveAntworten = (antworten || []).filter((a) => a && a.umfrageId);
        const next = saveState({
            umfragen: liveUmfragen,
            antworten: liveAntworten,
            localDemoOnly: false,
            updatedAt: new Date().toISOString()
        });
        const listTitle =
            (ctx.listUmfragen && ctx.listUmfragen.title ? ctx.listUmfragen.title : '') +
            (ctx.listAntworten && ctx.listAntworten.title
                ? ' + ' + ctx.listAntworten.title
                : '');
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            siteUrl: ctx.webUrl,
            listTitle,
            lastPullAt: new Date().toISOString()
        });
        return { ok: true, state: next, meta: getKollegiumSyncMeta() };
    } catch (e) {
        const denied = isKollegiumAccessDeniedError(e);
        const msg = String((e && e.message) || e || 'SharePoint nicht erreichbar.');
        let state = loadState();
        if (opts.forceDemoIfEmpty !== false && !(state.umfragen || []).length) {
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
            meta: getKollegiumSyncMeta(),
            softError: denied
                ? 'Kein SharePoint-Zugriff – lokal weiterarbeiten.'
                : msg
        };
    }
}

function markLiveOk(ctx) {
    setMeta({
        ok: true,
        connected: true,
        source: 'live',
        error: '',
        siteUrl: ctx.webUrl
    });
    const state = loadState();
    if (state.localDemoOnly) {
        state.localDemoOnly = false;
        saveState(state);
    }
}

/**
 * @param {object} umfrage
 */
export async function pushKollegiumUmfrage(umfrage) {
    if (!umfrage || isDemoSeedItem(umfrage)) {
        return { ok: true, skipped: true, item: umfrage };
    }
    try {
        const ctx = await ensureKollegiumContext();
        if (!ctx) {
            return {
                ok: false,
                item: umfrage,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (umfrage.spoItemId) {
            remote = await updateUmfrageItem(ctx, umfrage);
        } else {
            remote = await createUmfrageItem(ctx, umfrage);
        }
        const merged = {
            ...umfrage,
            id: umfrage.id || remote.id,
            spoItemId: remote.spoItemId || umfrage.spoItemId,
            fragen:
                Array.isArray(umfrage.fragen) && umfrage.fragen.length
                    ? umfrage.fragen
                    : remote.fragen
        };
        const res = upsertUmfrage(merged);
        markLiveOk(ctx);
        return { ok: true, item: res.item || merged, state: res.state };
    } catch (e) {
        const denied = isKollegiumAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: umfrage,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * @param {object} antwort
 */
export async function pushKollegiumAntwort(antwort) {
    if (!antwort || isDemoSeedItem(antwort)) {
        return { ok: true, skipped: true, item: antwort };
    }
    try {
        const ctx = await ensureKollegiumContext();
        if (!ctx) {
            return {
                ok: false,
                item: antwort,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (antwort.spoItemId) {
            remote = await updateAntwortItem(ctx, antwort);
        } else {
            remote = await createAntwortItem(ctx, antwort);
        }
        const merged = {
            ...antwort,
            id: antwort.id || remote.id,
            spoItemId: remote.spoItemId || antwort.spoItemId,
            werte:
                antwort.werte && Object.keys(antwort.werte).length
                    ? antwort.werte
                    : remote.werte
        };
        const res = upsertAntwort(merged);
        markLiveOk(ctx);
        return { ok: true, item: res.item || merged, state: res.state };
    } catch (e) {
        const denied = isKollegiumAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: antwort,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * @param {object} umfrage
 * @param {object[]} [antworten]
 */
export async function pushKollegiumDeleteUmfrage(umfrage, antworten) {
    if (!umfrage || isDemoSeedItem(umfrage)) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureKollegiumContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        const related = Array.isArray(antworten) ? antworten : [];
        for (let i = 0; i < related.length; i++) {
            const a = related[i];
            if (!a || isDemoSeedItem(a) || !a.spoItemId) continue;
            await deleteAntwortItem(ctx, a.spoItemId);
        }
        if (umfrage.spoItemId) {
            await deleteUmfrageItem(ctx, umfrage.spoItemId);
        }
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isKollegiumAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}

/**
 * @param {object} antwort
 */
export async function pushKollegiumDeleteAntwort(antwort) {
    if (!antwort || isDemoSeedItem(antwort) || !antwort.spoItemId) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureKollegiumContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        await deleteAntwortItem(ctx, antwort.spoItemId);
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isKollegiumAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}
