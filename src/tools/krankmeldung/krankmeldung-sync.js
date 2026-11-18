/**
 * Live-Sync Krankmeldung ↔ SharePoint (local-first).
 * 403/Fehler wipe’t lokalen Cache und Demo nicht.
 * Public-Formular: ohne Graph → lokal/Remote; Admin Pull/Push = SoR.
 * Kiosk (gleiche Session): pending lokale Items vor Pull nach SPO pushen.
 */

import {
    createMeldungItem,
    deleteMeldungItem,
    loadAllMeldungen,
    resolveKrankmeldungContext,
    updateMeldungItem
} from './krankmeldung-graph.js';
import { isDemoSeedItem } from './krankmeldung-model.js';
import { resolveKrankmeldungSiteUrl } from './krankmeldung-settings.js';
import {
    loadState,
    mergeRemoteMeldung,
    saveState,
    seedDemoIfEmpty
} from './krankmeldung-store.js';

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

export function getKrankmeldungSyncMeta() {
    return { ...syncMeta };
}

export function isKrankmeldungAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|accessDenied|Authorization_RequestDenied/i.test(msg);
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getKrankmeldungSyncMeta();
}

/**
 * @param {{ force?: boolean }} [opts]
 */
export async function ensureKrankmeldungContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = resolveKrankmeldungSiteUrl();
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
    cachedCtx = await resolveKrankmeldungContext(siteUrl);
    setMeta({
        siteUrl: cachedCtx.webUrl,
        listTitle: (cachedCtx.list && cachedCtx.list.title) || '',
        connected: true
    });
    return cachedCtx;
}

/**
 * Lokale Meldung nach SPO pushen (Create oder Update).
 * @param {object} meldung
 */
export async function pushKrankmeldungMeldung(meldung) {
    if (!meldung || isDemoSeedItem(meldung)) {
        return { ok: true, skipped: true, item: meldung };
    }
    try {
        const ctx = await ensureKrankmeldungContext();
        if (!ctx) {
            return {
                ok: false,
                item: meldung,
                error: 'Keine SharePoint-Site – nur lokal gespeichert.'
            };
        }
        let remote;
        if (meldung.spoItemId) {
            remote = await updateMeldungItem(ctx, meldung);
        } else {
            remote = await createMeldungItem(ctx, meldung);
        }
        const merged = {
            ...meldung,
            ...remote,
            id: meldung.id || remote.id,
            spoItemId: remote.spoItemId || meldung.spoItemId
        };
        const res = mergeRemoteMeldung(merged);
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
        const denied = isKrankmeldungAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        setMeta({
            ok: false,
            connected: !denied,
            source: 'local',
            error: denied ? '' : msg
        });
        return {
            ok: false,
            item: meldung,
            error: denied
                ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
                : msg
        };
    }
}

/**
 * Pending lokale Items (ohne spoItemId, kein Demo) vor Pull nach SPO.
 */
export async function pushPendingKrankmeldungen() {
    const state = loadState();
    const pending = (state.meldungen || []).filter(
        (m) => m && !isDemoSeedItem(m) && !String(m.spoItemId || '').trim()
    );
    let pushed = 0;
    let err = '';
    for (const m of pending) {
        const r = await pushKrankmeldungMeldung(m);
        if (r.ok && !r.skipped) pushed++;
        else if (!r.ok && !r.skipped && !err) err = r.error || 'Push fehlgeschlagen.';
    }
    return { ok: !err || pushed > 0, pushed, pending: pending.length, error: err };
}

/**
 * Pull von SharePoint. Bei Fehler: lokaler Stand bleibt.
 * @param {{ forceDemoIfEmpty?: boolean, pushPendingFirst?: boolean }} [opts]
 */
export async function pullKrankmeldungFromSharePoint(opts = {}) {
    const siteUrl = resolveKrankmeldungSiteUrl();
    if (!siteUrl) {
        const state = seedDemoIfEmpty();
        setMeta({
            ok: !!state.localDemoOnly,
            connected: false,
            source: state.localDemoOnly ? 'demo' : 'local',
            error: 'Keine Site-URL hinterlegt.',
            siteUrl: ''
        });
        return { ok: false, state, meta: getKrankmeldungSyncMeta() };
    }

    if (opts.pushPendingFirst !== false) {
        try {
            await pushPendingKrankmeldungen();
        } catch {
            /* soft – Pull versucht trotzdem */
        }
    }

    try {
        const { ctx, items } = await loadAllMeldungen({ webUrl: siteUrl });
        cachedCtx = ctx;
        const live = (items || []).filter((m) => m && (m.schuelerName || m.id));
        const next = saveState({
            meldungen: live,
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
        return { ok: true, state: next, meta: getKrankmeldungSyncMeta() };
    } catch (e) {
        const denied = isKrankmeldungAccessDeniedError(e);
        const msg = String((e && e.message) || e || 'SharePoint nicht erreichbar.');
        let state = loadState();
        if (opts.forceDemoIfEmpty !== false && !(state.meldungen || []).length) {
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
            meta: getKrankmeldungSyncMeta(),
            softError: denied
                ? 'Kein SharePoint-Zugriff – lokal weiterarbeiten.'
                : msg
        };
    }
}

/**
 * @param {object} meldung
 */
export async function pushKrankmeldungDelete(meldung) {
    if (!meldung || isDemoSeedItem(meldung) || !meldung.spoItemId) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureKrankmeldungContext();
        if (!ctx) return { ok: false, error: 'Keine SharePoint-Site.' };
        await deleteMeldungItem(ctx, meldung.spoItemId);
        setMeta({ ok: true, connected: true, source: 'live', error: '' });
        return { ok: true };
    } catch (e) {
        const denied = isKrankmeldungAccessDeniedError(e);
        const msg = String((e && e.message) || e || '');
        return {
            ok: false,
            error: denied ? 'Löschen in SharePoint verweigert.' : msg
        };
    }
}
