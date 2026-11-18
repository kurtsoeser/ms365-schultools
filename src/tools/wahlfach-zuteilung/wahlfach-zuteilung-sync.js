/**
 * Soft Live-Sync Wahlfach-Zuteilung ↔ SharePoint (4 Listen).
 * Demo-Seeds werden nicht gepusht; 403 wipe’t lokalen Stand nicht.
 */

import { DEMO_SEED_TAG, isWzDemoSeedState } from './wahlfach-zuteilung-demo.js';
import {
    deleteWzAngebotItem,
    deleteWzZuteilungItem,
    pushWzAnmeldung,
    resolveWzContext,
    softPullWzFromSharePoint,
    upsertWzAngebotItem,
    upsertWzRundeItem,
    upsertWzZuteilungItem
} from './wahlfach-zuteilung-graph.js';
import { isWzAccessDeniedError } from './wahlfach-zuteilung-logic.js';
import { resolveWzSiteUrl } from './wahlfach-zuteilung-state.js';
import {
    hasLocalWzData,
    loadState,
    saveState,
    upsertAnmeldung,
    upsertAngebot,
    upsertRunde
} from './wahlfach-zuteilung-store.js';

/** @type {{
 *   ok: boolean,
 *   source: 'live'|'local'|'demo'|'none',
 *   error: string,
 *   softError: string,
 *   siteUrl: string,
 *   lastPullAt: string,
 *   lastPushAt: string,
 *   connected: boolean
 * }} */
let syncMeta = {
    ok: false,
    source: 'none',
    error: '',
    softError: '',
    siteUrl: '',
    lastPullAt: '',
    lastPushAt: '',
    connected: false
};

/** @type {object|null} */
let cachedCtx = null;

export function getWzSyncMeta() {
    return { ...syncMeta };
}

function setMeta(partial) {
    syncMeta = { ...syncMeta, ...partial };
    return getWzSyncMeta();
}

/**
 * @param {object|null|undefined} entity
 * @returns {boolean}
 */
export function isWzDemoEntity(entity) {
    if (!entity) return true;
    if (entity.seedTag && String(entity.seedTag) === DEMO_SEED_TAG) return true;
    const id = String(entity.id || entity.anmeldungId || '');
    return (
        id.indexOf('runde-demo-') === 0 ||
        id.indexOf('ang-demo-') === 0 ||
        id.indexOf('anm-demo-') === 0
    );
}

/**
 * @param {{ force?: boolean, siteUrl?: string }} [opts]
 */
export async function ensureWzContext(opts = {}) {
    if (cachedCtx && !opts.force) return cachedCtx;
    const siteUrl = String(opts.siteUrl || resolveWzSiteUrl() || '').trim();
    if (!siteUrl) {
        cachedCtx = null;
        setMeta({
            ok: false,
            connected: false,
            source: syncMeta.source === 'demo' ? 'demo' : 'none',
            softError: 'Keine SharePoint-Site – lokal weiterarbeiten.',
            siteUrl: ''
        });
        return null;
    }
    cachedCtx = await resolveWzContext(siteUrl);
    setMeta({
        siteUrl: cachedCtx.webUrl,
        connected: true
    });
    return cachedCtx;
}

function pushDeniedMessage(e) {
    return isWzAccessDeniedError(e)
        ? 'Kein Schreibzugriff auf SharePoint – lokal gespeichert.'
        : String((e && e.message) || e || 'SharePoint-Sync fehlgeschlagen.');
}

/**
 * Soft-Pull. Demo/Local bleiben bei Fehler.
 * @param {{ localDemoOnly?: boolean, siteUrl?: string }} [opts]
 */
export async function softPullWz(opts = {}) {
    const siteUrl = String(opts.siteUrl || resolveWzSiteUrl() || '').trim();
    const local = loadState();
    const pull = await softPullWzFromSharePoint(siteUrl, {
        localDemoOnly: !!opts.localDemoOnly,
        hasLocalData: hasLocalWzData(local) || isWzDemoSeedState(local)
    });
    if (pull.ok && pull.state) {
        const next = saveState(pull.state);
        cachedCtx = null;
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            error: '',
            softError: '',
            siteUrl,
            lastPullAt: new Date().toISOString()
        });
        return { ok: true, state: next, siteId: pull.siteId, meta: getWzSyncMeta() };
    }
    setMeta({
        ok: false,
        connected: false,
        source: opts.localDemoOnly || isWzDemoSeedState(local) ? 'demo' : 'local',
        softError: pull.softError || '',
        error: pull.error || '',
        siteUrl
    });
    return {
        ok: false,
        state: local,
        softError: pull.softError || '',
        error: pull.error || '',
        siteId: '',
        meta: getWzSyncMeta()
    };
}

/**
 * @param {object} runde
 * @param {{ localDemoOnly?: boolean }} [opts]
 */
export async function softPushWzRunde(runde, opts = {}) {
    if (opts.localDemoOnly || isWzDemoEntity(runde) || isWzDemoSeedState(loadState())) {
        return { ok: true, skipped: true, item: runde };
    }
    try {
        const ctx = await ensureWzContext();
        if (!ctx) {
            return { ok: false, item: runde, softError: syncMeta.softError || 'Keine SharePoint-Site.' };
        }
        const remote = await upsertWzRundeItem(ctx, runde);
        const res = upsertRunde(remote);
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            softError: '',
            error: '',
            lastPushAt: new Date().toISOString()
        });
        return { ok: true, item: remote, state: res.state };
    } catch (e) {
        const softError = pushDeniedMessage(e);
        setMeta({ ok: false, connected: !isWzAccessDeniedError(e), source: 'local', softError });
        return { ok: false, item: runde, softError };
    }
}

/**
 * @param {object} angebot
 * @param {{ localDemoOnly?: boolean }} [opts]
 */
export async function softPushWzAngebot(angebot, opts = {}) {
    if (opts.localDemoOnly || isWzDemoEntity(angebot) || isWzDemoSeedState(loadState())) {
        return { ok: true, skipped: true, item: angebot };
    }
    try {
        const ctx = await ensureWzContext();
        if (!ctx) {
            return {
                ok: false,
                item: angebot,
                softError: syncMeta.softError || 'Keine SharePoint-Site.'
            };
        }
        const state = loadState();
        const remote = await upsertWzAngebotItem(ctx, angebot, { runde: state.runde });
        const res = upsertAngebot(remote);
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            softError: '',
            error: '',
            lastPushAt: new Date().toISOString()
        });
        return { ok: true, item: remote, state: res.state };
    } catch (e) {
        const softError = pushDeniedMessage(e);
        setMeta({ ok: false, connected: !isWzAccessDeniedError(e), source: 'local', softError });
        return { ok: false, item: angebot, softError };
    }
}

/**
 * @param {object} angebot
 * @param {{ localDemoOnly?: boolean }} [opts]
 */
export async function softPushWzAngebotDelete(angebot, opts = {}) {
    if (
        opts.localDemoOnly ||
        !angebot ||
        isWzDemoEntity(angebot) ||
        !angebot.spoItemId ||
        isWzDemoSeedState(loadState())
    ) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureWzContext();
        if (!ctx) return { ok: false, softError: 'Keine SharePoint-Site.' };
        await deleteWzAngebotItem(ctx, angebot.spoItemId);
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            softError: '',
            lastPushAt: new Date().toISOString()
        });
        return { ok: true };
    } catch (e) {
        const softError = isWzAccessDeniedError(e)
            ? 'Löschen in SharePoint verweigert – lokal entfernt.'
            : pushDeniedMessage(e);
        setMeta({ ok: false, source: 'local', softError });
        return { ok: false, softError };
    }
}

/**
 * @param {object} anmeldung
 * @param {{ localDemoOnly?: boolean }} [opts]
 */
export async function softPushWzAnmeldung(anmeldung, opts = {}) {
    if (opts.localDemoOnly || isWzDemoEntity(anmeldung) || isWzDemoSeedState(loadState())) {
        return { ok: true, skipped: true, item: anmeldung };
    }
    try {
        const ctx = await ensureWzContext();
        if (!ctx) {
            return {
                ok: false,
                item: anmeldung,
                softError: syncMeta.softError || 'Keine SharePoint-Site.'
            };
        }
        const state = loadState();
        const remote = await pushWzAnmeldung(ctx, anmeldung, {
            runde: state.runde,
            angebote: state.angebote
        });
        const res = upsertAnmeldung(remote);
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            softError: '',
            error: '',
            lastPushAt: new Date().toISOString()
        });
        return { ok: true, item: remote, state: res.state };
    } catch (e) {
        const softError = pushDeniedMessage(e);
        setMeta({ ok: false, connected: !isWzAccessDeniedError(e), source: 'local', softError });
        return { ok: false, item: anmeldung, softError };
    }
}

/**
 * Zuteilung pushen oder bei Aufheben SPO-Item löschen.
 * @param {{ anmeldungId: string, angebotId: string|null, prevZuteilung?: object|null, zuteilung?: object|null }} payload
 * @param {{ localDemoOnly?: boolean }} [opts]
 */
export async function softPushWzZuteilung(payload, opts = {}) {
    const p = payload || {};
    const state = loadState();
    if (opts.localDemoOnly || isWzDemoSeedState(state)) {
        return { ok: true, skipped: true };
    }
    try {
        const ctx = await ensureWzContext();
        if (!ctx) {
            return { ok: false, softError: syncMeta.softError || 'Keine SharePoint-Site.' };
        }
        const prev = p.prevZuteilung || null;
        if (!p.angebotId) {
            const spoId = (prev && prev.spoItemId) || '';
            if (spoId) await deleteWzZuteilungItem(ctx, spoId);
            setMeta({
                ok: true,
                connected: true,
                source: 'live',
                softError: '',
                lastPushAt: new Date().toISOString()
            });
            return { ok: true, deleted: true };
        }
        const z = p.zuteilung || state.zuteilungen.find((x) => x && x.anmeldungId === p.anmeldungId);
        if (!z) return { ok: false, softError: 'Zuteilung lokal nicht gefunden.' };
        const remote = await upsertWzZuteilungItem(ctx, z, {
            runde: state.runde,
            angebote: state.angebote,
            anmeldungen: state.anmeldungen
        });
        const next = loadState();
        next.zuteilungen = (next.zuteilungen || []).map((row) =>
            row && row.anmeldungId === remote.anmeldungId ? { ...row, spoItemId: remote.spoItemId } : row
        );
        saveState(next);
        setMeta({
            ok: true,
            connected: true,
            source: 'live',
            softError: '',
            lastPushAt: new Date().toISOString()
        });
        return { ok: true, item: remote, state: next };
    } catch (e) {
        const softError = pushDeniedMessage(e);
        setMeta({ ok: false, connected: !isWzAccessDeniedError(e), source: 'local', softError });
        return { ok: false, softError };
    }
}
