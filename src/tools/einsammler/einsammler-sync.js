/**
 * Live-Sync Einsammler ↔ SharePoint (local-first).
 * 403/Fehler wipe’t lokalen Cache und Demo nicht.
 * Public-Antworten: ohne Graph – Inbox/Code/Remote; Absender-Push nach Übernahme.
 */

import {
    createAntwortItem,
    createKopfItem,
    deleteAntwortItem,
    deleteKopfItem,
    loadAllEinsammler,
    resolveEinsammlerContext,
    updateAntwortItem,
    updateKopfItem
} from './einsammler-graph.js';
import { isDemoSeedItem } from './einsammler-model.js';
import { resolveEinsammlerSiteUrl } from './einsammler-settings.js';
import {
    loadState,
    saveState,
    seedDemoIfEmpty,
    upsertAntwort,
    upsertEinsammler
} from './einsammler-store.js';

/** @type {{
 *   ok: boolean,
 *   source: 'live'|'local'|'demo'|'none',
 *   error: string,
 *   siteUrl: string,
 *   listTitle: string,
 *   listAntwortenTitle: string,
 *   lastPullAt: string,
 *   connected: boolean
 * }} */
let syncMeta = {
    ok: false,
    source: 'none',
    error: '',
    siteUrl: '',
    listTitle: '',
    listAntwortenTitle: '',
    lastPullAt: '',
    connected: false
};

/** @type {object|null} */
let cachedCtx = null;

export function getEinsammlerSyncMeta() {
    return { ...syncMeta };
}

export function isEinsammlerAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|accessDenied|Authorization_RequestDenied/i.test(msg);
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getEinsammlerSyncMeta();
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function ensureEinsammlerContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = resolveEinsammlerSiteUrl();
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
    cachedCtx = await resolveEinsammlerContext(siteUrl);
    setMeta({
        siteUrl: cachedCtx.webUrl,
        listTitle: (cachedCtx.listKoepfe && cachedCtx.listKoepfe.title) || '',
        listAntwortenTitle: (cachedCtx.listAntworten && cachedCtx.listAntworten.title) || '',
        connected: true
    });
    return cachedCtx;
}

/**
 * Pull von SharePoint. Bei Fehler: lokaler Stand bleibt.
 * @param {{ forceDemoIfEmpty?: boolean }} [opts]
 */
export async function pullEinsammlerFromSharePoint(opts = {}) {
    const siteUrl = resolveEinsammlerSiteUrl();
    if (!siteUrl) {
        const state = seedDemoIfEmpty();
        setMeta({
            ok: !!state.localDemoOnly,
            connected: false,
            source: state.localDemoOnly ? 'demo' : 'local',
            error: 'Keine Site-URL hinterlegt.',
            siteUrl: ''
        });
        return { ok: false, state, meta: getEinsammlerSyncMeta() };
    }

    try {
        const { ctx, einsammler, antworten } = await loadAllEinsammler({ webUrl: siteUrl });
        cachedCtx = ctx;
        const liveHeads = (einsammler || []).filter((m) => m && (m.titel || m.id));
        const liveAnt = (antworten || []).filter((a) => a && (a.id || a.einsammlerId));
        const next = saveState({
            einsammler: liveHeads,
            antworten: liveAnt,
            localDemoOnly: false,
            updatedAt: new Date().toISOString()
        });
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            siteUrl: ctx.webUrl,
            listTitle: (ctx.listKoepfe && ctx.listKoepfe.title) || '',
            listAntwortenTitle: (ctx.listAntworten && ctx.listAntworten.title) || '',
            lastPullAt: new Date().toISOString()
        });
        return { ok: true, state: next, meta: getEinsammlerSyncMeta() };
    } catch (e) {
        const denied = isEinsammlerAccessDeniedError(e);
        const msg = String((e && e.message) || e || 'SharePoint nicht erreichbar.');
        let state = loadState();
        if (
            opts.forceDemoIfEmpty !== false &&
            !(state.einsammler || []).length &&
            !(state.antworten || []).length
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
            meta: getEinsammlerSyncMeta(),
            softError: denied
                ? 'Kein SharePoint-Zugriff – lokal weiterarbeiten.'
                : msg
        };
    }
}

/**
 * @param {object} kopf
 */
export async function pushEinsammlerKopf(kopf) {
    if (!kopf || isDemoSeedItem(kopf)) {
        return { ok: true, skipped: true, item: kopf };
    }
    try {
        const ctx = await ensureEinsammlerContext();
        if (!ctx) {
            return {
                ok: false,
                item: kopf,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (kopf.spoItemId) {
            remote = await updateKopfItem(ctx, kopf);
        } else {
            remote = await createKopfItem(ctx, kopf);
        }
        const merged = {
            ...kopf,
            ...remote,
            id: kopf.id || remote.id,
            spoItemId: remote.spoItemId || kopf.spoItemId
        };
        const res = upsertEinsammler(merged);
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
        const denied = isEinsammlerAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: kopf,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * @param {object} antwort
 */
export async function pushEinsammlerAntwort(antwort) {
    if (!antwort || isDemoSeedItem(antwort)) {
        return { ok: true, skipped: true, item: antwort };
    }
    try {
        const ctx = await ensureEinsammlerContext();
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
            ...remote,
            id: antwort.id || remote.id,
            spoItemId: remote.spoItemId || antwort.spoItemId
        };
        const res = upsertAntwort(merged);
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
        const denied = isEinsammlerAccessDeniedError(e);
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
 * @param {object[]} antworten
 */
export async function pushEinsammlerAntworten(antworten) {
    const list = Array.isArray(antworten) ? antworten : [];
    /** @type {object[]} */
    const results = [];
    let okCount = 0;
    let err = '';
    for (const a of list) {
        const r = await pushEinsammlerAntwort(a);
        results.push(r);
        if (r.ok || r.skipped) okCount++;
        else if (!err) err = r.error || 'Push fehlgeschlagen.';
    }
    return {
        ok: okCount === list.length,
        results,
        error: err,
        pushed: okCount
    };
}

/**
 * @param {object} kopf
 * @param {object[]} [antworten]
 */
export async function pushEinsammlerDelete(kopf, antworten) {
    if (kopf && isDemoSeedItem(kopf)) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureEinsammlerContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        const ants = Array.isArray(antworten) ? antworten : [];
        for (const a of ants) {
            if (!a || isDemoSeedItem(a) || !a.spoItemId) continue;
            await deleteAntwortItem(ctx, a.spoItemId);
        }
        if (kopf && kopf.spoItemId) {
            await deleteKopfItem(ctx, kopf.spoItemId);
        }
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isEinsammlerAccessDeniedError(e);
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
export async function pushEinsammlerAntwortDelete(antwort) {
    if (!antwort || isDemoSeedItem(antwort) || !antwort.spoItemId) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureEinsammlerContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        await deleteAntwortItem(ctx, antwort.spoItemId);
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isEinsammlerAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}
