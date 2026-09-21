/**
 * Automatischer Stammdaten-Abgleich mit SharePoint (IT-Bibliothek).
 *
 * Phase 1: einmaliger Pull nach Login (pro Tab-Session)
 * Phase 2: debounced Push nach lokalen Änderungen + Flush beim Verlassen
 * Phase 3: Versionsvergleich (lastModified / exportedAt) + Status-Events
 *
 * Voraussetzung: IT-Bibliothek eingerichtet, User angemeldet, kein Demo-Modus.
 */
import {
    DEFAULT_FOLDER,
    IT_LIBRARY_TITLE,
    isAutoSyncIgnoredChangeSource,
    shouldApplyRemoteBackup,
    formatSyncStatusDe
} from './stammdaten-sharepoint-sync-logic.js';
import {
    isReady,
    loadItMeta,
    loadLocalSyncMeta,
    saveLocalSyncMeta,
    setLocalDirty,
    isLocalDirty,
    getCurrentBackupRemoteInfo,
    downloadCurrentBackup,
    uploadCurrentBackup
} from './stammdaten-sharepoint-sync-api.js';

const SESSION_PULL_KEY = 'ms365-spo-auto-pull-done-v1';
const PUSH_DEBOUNCE_MS = 3000;
const FOLDER = DEFAULT_FOLDER;

/** @type {{ phase: string, error: string, lastPullAt: string, lastPushAt: string, message: string }} */
let live = {
    phase: 'idle',
    error: '',
    lastPullAt: '',
    lastPushAt: '',
    message: ''
};

let started = false;
let pushTimer = null;
let pushInFlight = null;
let pullInFlight = null;
let suppressPushUntil = 0;

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
}

function isLoggedIn() {
    try {
        return typeof window.ms365AuthIsLoggedIn === 'function' && !!window.ms365AuthIsLoggedIn();
    } catch {
        return false;
    }
}

function isDemoActive() {
    try {
        return !!(window.ms365DemoMode && typeof window.ms365DemoMode.isActive === 'function' && window.ms365DemoMode.isActive());
    } catch {
        return false;
    }
}

function sessionPullDone() {
    try {
        return sessionStorage.getItem(SESSION_PULL_KEY) === '1';
    } catch {
        return false;
    }
}

function markSessionPullDone() {
    try {
        sessionStorage.setItem(SESSION_PULL_KEY, '1');
    } catch {
        /* ignore */
    }
}

function setLive(partial) {
    live = Object.assign({}, live, partial || {});
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-spo-sync-status', {
                detail: getStatus()
            })
        );
    } catch {
        /* ignore */
    }
}

export function getStatus() {
    const meta = loadLocalSyncMeta();
    const it = loadItMeta();
    const ready = isReady();
    const dirty = !!(meta && meta.dirty) || live.phase === 'pushing';
    const lastAt = meta && meta.at ? meta.at : live.lastPushAt || live.lastPullAt || '';
    const lastDirection = (meta && meta.direction) || '';
    return {
        ready: ready,
        phase: live.phase,
        dirty: !!dirty,
        error: live.error || (meta && meta.pendingError) || '',
        lastAt: lastAt,
        lastDirection: lastDirection,
        lastPullAt: live.lastPullAt,
        lastPushAt: live.lastPushAt,
        libraryTitle: (it && it.listTitle) || IT_LIBRARY_TITLE,
        message: live.message || formatSyncStatusDe({
            ready: ready,
            phase: live.phase,
            dirty: !!dirty,
            lastAt: lastAt,
            lastDirection: lastDirection,
            error: live.error,
            libraryTitle: (it && it.listTitle) || IT_LIBRARY_TITLE
        })
    };
}

function canRunNetworkSync() {
    return isReady() && isLoggedIn() && !isDemoActive();
}

/**
 * Phase 1: SharePoint → lokal (einmal pro Tab-Session).
 * @param {{ force?: boolean, reloadOnApply?: boolean }} [opts]
 */
export async function runSessionPull(opts) {
    const options = opts || {};
    if (!options.force && sessionPullDone()) return { skipped: true, reason: 'session-done' };
    if (!canRunNetworkSync()) return { skipped: true, reason: 'not-ready' };
    if (pullInFlight) return pullInFlight;

    pullInFlight = (async function () {
        setLive({ phase: 'pulling', error: '', message: 'Lade Stammdaten von SharePoint …' });
        try {
            if (isLocalDirty()) {
                // Lokale Änderungen nicht überschreiben – zuerst pushen.
                markSessionPullDone();
                setLive({ phase: 'idle', message: '' });
                schedulePush(0);
                return { skipped: true, reason: 'local-dirty' };
            }

            const remote = await getCurrentBackupRemoteInfo({ folder: FOLDER });
            const localMeta = loadLocalSyncMeta();
            const decision = shouldApplyRemoteBackup({
                remoteExists: !!(remote && remote.exists),
                remoteLastModified: remote && remote.lastModifiedDateTime,
                remoteExportedAt: '',
                localDirty: false,
                localRemoteLastModified: localMeta && localMeta.remoteLastModified,
                localRemoteExportedAt: localMeta && localMeta.remoteExportedAt
            });

            if (!remote.exists) {
                markSessionPullDone();
                setLive({ phase: 'idle', message: 'Noch kein Backup auf SharePoint' });
                return { skipped: true, reason: 'missing' };
            }

            if (!decision.apply && decision.reason !== 'never-synced') {
                // Bei same-modified fertig; bei Unsicherheit trotzdem Payload prüfen.
                if (decision.reason === 'same-modified' || decision.reason === 'local-current') {
                    markSessionPullDone();
                    setLive({ phase: 'idle', message: '' });
                    return { skipped: true, reason: decision.reason };
                }
            }

            const preview = await downloadCurrentBackup({
                folder: FOLDER,
                apply: false,
                recordMeta: false
            });
            const payload = preview.payload || {};
            const item = preview.item || {};
            const decision2 = shouldApplyRemoteBackup({
                remoteExists: true,
                remoteLastModified: item.lastModifiedDateTime || remote.lastModifiedDateTime,
                remoteExportedAt: payload.exportedAt || '',
                localDirty: false,
                localRemoteLastModified: localMeta && localMeta.remoteLastModified,
                localRemoteExportedAt: localMeta && localMeta.remoteExportedAt
            });

            const syncMetaFromRemote = {
                at: new Date().toISOString(),
                direction: 'pull',
                fileName: item.name || 'ms365-stammdaten-aktuell.json',
                folder: FOLDER,
                webUrl: item.webUrl || remote.webUrl || '',
                siteUrl: (loadItMeta() && loadItMeta().siteUrl) || '',
                driveId: (loadItMeta() && loadItMeta().driveId) || '',
                summary: payload.schoolName || payload.domain || '',
                remoteETag: item.eTag || item.cTag || remote.eTag || '',
                remoteLastModified: item.lastModifiedDateTime || remote.lastModifiedDateTime || '',
                remoteExportedAt: payload.exportedAt || '',
                dirty: false,
                pendingError: null
            };

            if (!decision2.apply) {
                saveLocalSyncMeta(
                    Object.assign({}, loadLocalSyncMeta(), {
                        remoteETag: syncMetaFromRemote.remoteETag,
                        remoteLastModified: syncMetaFromRemote.remoteLastModified,
                        remoteExportedAt: syncMetaFromRemote.remoteExportedAt,
                        dirty: false
                    })
                );
                markSessionPullDone();
                setLive({ phase: 'idle', message: '' });
                return { skipped: true, reason: decision2.reason };
            }

            suppressPushUntil = Date.now() + 5000;
            const bb = window.ms365BrowserBackup;
            if (!bb || typeof bb.importPayload !== 'function') {
                throw new Error('Backup-Modul fehlt.');
            }
            markSessionPullDone();
            bb.importPayload(payload);
            saveLocalSyncMeta(syncMetaFromRemote);

            live.lastPullAt = syncMetaFromRemote.at;
            setLive({ phase: 'idle', error: '', message: 'Stammdaten von SharePoint aktualisiert' });
            toast('Stammdaten von SharePoint aktualisiert.');

            const reload = options.reloadOnApply !== false;
            if (reload) {
                try {
                    window.dispatchEvent(
                        new CustomEvent('ms365-tenant-settings-changed', {
                            detail: { source: 'spo-auto-pull' }
                        })
                    );
                } catch {
                    /* ignore */
                }
                window.setTimeout(function () {
                    window.location.reload();
                }, 200);
            }
            return { applied: true, payload: payload };
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            // Einmal pro Session markieren, damit Login-Seiten nicht in einer Fehler-Schleife landen
            markSessionPullDone();
            setLive({ phase: 'error', error: msg, message: msg });
            return { error: msg };
        } finally {
            pullInFlight = null;
            if (live.phase === 'pulling') setLive({ phase: 'idle' });
        }
    })();

    return pullInFlight;
}

/**
 * Phase 2: lokal → SharePoint.
 * @param {{ keepDated?: boolean, reason?: string }} [opts]
 */
export async function runPush(opts) {
    const options = opts || {};
    if (!canRunNetworkSync()) return { skipped: true, reason: 'not-ready' };
    if (Date.now() < suppressPushUntil) return { skipped: true, reason: 'suppressed' };
    if (pushInFlight) return pushInFlight;

    pushInFlight = (async function () {
        setLive({ phase: 'pushing', error: '', message: 'Sichere Stammdaten nach SharePoint …' });
        try {
            const result = await uploadCurrentBackup({
                folder: FOLDER,
                keepDated: !!options.keepDated
            });
            live.lastPushAt = new Date().toISOString();
            setLive({ phase: 'idle', error: '', message: '' });
            return { ok: true, result: result };
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            try {
                const cur = loadLocalSyncMeta();
                saveLocalSyncMeta(Object.assign({}, cur, { dirty: true, pendingError: msg }));
            } catch {
                /* ignore */
            }
            setLive({ phase: 'error', error: msg, message: msg });
            return { error: msg };
        } finally {
            pushInFlight = null;
            if (live.phase === 'pushing') setLive({ phase: 'idle' });
        }
    })();

    return pushInFlight;
}

export function schedulePush(delayMs) {
    if (!canRunNetworkSync()) return;
    if (Date.now() < suppressPushUntil) return;
    const ms = delayMs == null ? PUSH_DEBOUNCE_MS : Math.max(0, Number(delayMs) || 0);
    if (pushTimer) {
        clearTimeout(pushTimer);
        pushTimer = null;
    }
    setLocalDirty(true);
    setLive({ message: '' });
    pushTimer = setTimeout(function () {
        pushTimer = null;
        runPush({ reason: 'debounce' }).then(function () {
            try {
                window.dispatchEvent(new CustomEvent('ms365-spo-sync-status', { detail: getStatus() }));
            } catch {
                /* ignore */
            }
        });
    }, ms);
}

/** Offene Änderungen sofort schreiben (Tab-Wechsel / Schließen). */
export function flushPush() {
    if (pushTimer) {
        clearTimeout(pushTimer);
        pushTimer = null;
    }
    if (!canRunNetworkSync()) return Promise.resolve({ skipped: true });
    if (!isLocalDirty() && live.phase !== 'pushing') return Promise.resolve({ skipped: true });
    return runPush({ reason: 'flush' });
}

function onTenantSettingsChanged(ev) {
    const detail = (ev && ev.detail) || {};
    const src = detail.source || detail.reason || '';
    if (isAutoSyncIgnoredChangeSource(src)) return;
    if (!canRunNetworkSync()) return;
    schedulePush(PUSH_DEBOUNCE_MS);
}

function onAuthReady() {
    if (!isLoggedIn()) return;
    runSessionPull({ reloadOnApply: true }).then(function () {
        try {
            window.dispatchEvent(new CustomEvent('ms365-spo-sync-status', { detail: getStatus() }));
        } catch {
            /* ignore */
        }
    });
}

function onVisibilityFlush() {
    if (document.visibilityState === 'hidden') {
        flushPush();
    }
}

/**
 * Listener + einmaliger Session-Pull starten (idempotent).
 */
export function start() {
    if (started) return getStatus();
    started = true;

    window.addEventListener('ms365-tenant-settings-changed', onTenantSettingsChanged);
    window.addEventListener('ms365-auth-widget-ready', onAuthReady);
    window.addEventListener('ms365-auth-state-changed', onAuthReady);
    document.addEventListener('visibilitychange', onVisibilityFlush);
    window.addEventListener('pagehide', function () {
        flushPush();
    });

    // Falls Auth schon da ist
    if (isLoggedIn()) {
        onAuthReady();
    } else {
        // kurze Nachversuche nach MSAL-Init
        let tries = 0;
        const t = setInterval(function () {
            tries += 1;
            if (isLoggedIn() || tries > 20) {
                clearInterval(t);
                if (isLoggedIn()) onAuthReady();
            }
        }, 500);
    }

    return getStatus();
}

export default {
    start: start,
    getStatus: getStatus,
    runSessionPull: runSessionPull,
    runPush: runPush,
    schedulePush: schedulePush,
    flushPush: flushPush
};

// Auto-Start, wenn als Modul geladen
start();

window.ms365StammdatenSpoAutoSync = {
    start: start,
    getStatus: getStatus,
    runSessionPull: runSessionPull,
    runPush: runPush,
    schedulePush: schedulePush,
    flushPush: flushPush
};
