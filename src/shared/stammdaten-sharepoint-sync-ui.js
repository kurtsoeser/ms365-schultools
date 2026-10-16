/**
 * Buttons: Stammdaten ↔ SharePoint IT-Bibliothek (Upload / Laden).
 * data-ms365-spo-sync="upload" | "load" | "setup"
 * Optional: #stammdatenSpoSyncStatus für Kurzstatus.
 * Startet außerdem den Auto-Sync (Pull nach Login, Push nach Änderungen).
 *
 * Lokal-First: Upload/Load/Setup deaktiviert; Auto-Sync startet nur als No-Op.
 */
import {
    isSpoNetworkSyncAllowed,
    localFirstStatusMessage,
    applyLocalFirstSpoUi
} from './stammdaten-local-first.js';
import { DEFAULT_FOLDER, IT_LIBRARY_TITLE, formatSyncStatusDe } from './stammdaten-sharepoint-sync-logic.js';
import {
    isReady,
    setupPageHref,
    loadItMeta,
    loadLocalSyncMeta,
    uploadCurrentBackup
} from './stammdaten-sharepoint-sync-api.js';
import * as autoSync from './stammdaten-sharepoint-auto-sync.js';
import { resolveSharePointImportChoice } from './stammdaten-sharepoint-import-prompt.js';
import {
    applySharePointBackupLocally,
    applySharePointSelectiveImportLocally
} from './stammdaten-sharepoint-pull-apply.js';
import { pickAndDownloadSharePointBackup } from './stammdaten-sharepoint-backup-picker-ui.js';

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function confirmAsync(msg, opts) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, opts || {});
    }
    return Promise.resolve(window.confirm(msg));
}

function goToSetup(reason) {
    const href = setupPageHref();
    if (reason) toast(reason);
    window.location.href = href;
}

function setBusy(btns, busy) {
    (btns || []).forEach(function (btn) {
        if (!btn) return;
        btn.disabled = !!busy;
        if (busy) btn.setAttribute('aria-busy', 'true');
        else btn.removeAttribute('aria-busy');
    });
}

function refreshStatus() {
    const el = document.getElementById('stammdatenSpoSyncStatus');
    if (!el) return;
    if (!isSpoNetworkSyncAllowed()) {
        el.textContent = localFirstStatusMessage();
        return;
    }
    try {
        if (autoSync && typeof autoSync.getStatus === 'function') {
            const st = autoSync.getStatus();
            el.textContent = st.message || formatSyncStatusDe(st);
            return;
        }
    } catch {
        /* fallback below */
    }
    const it = loadItMeta();
    const m = loadLocalSyncMeta();
    el.textContent = formatSyncStatusDe({
        ready: isReady(),
        phase: 'idle',
        dirty: !!(m && m.dirty),
        lastAt: m && m.at,
        lastDirection: m && m.direction,
        error: m && m.pendingError,
        libraryTitle: (it && it.listTitle) || IT_LIBRARY_TITLE
    });
}

async function runUpload(btns) {
    if (!isSpoNetworkSyncAllowed()) {
        toast(localFirstStatusMessage());
        return;
    }
    if (!isReady()) {
        goToSetup('Zuerst die IT-Sicherungsbibliothek einrichten.');
        return;
    }
    const ok = await confirmAsync(
        'Aktuelle Schul- und App-Daten (vollständiges Browser-Backup inkl. Stammdaten, Schul-Organigramm, Power Automate und Werkzeugstände) jetzt in die SharePoint-IT-Bibliothek „' +
            (loadItMeta().listTitle || IT_LIBRARY_TITLE) +
            '“ schreiben?',
        { title: 'Nach SharePoint sichern', confirmLabel: 'Sichern', cancelLabel: 'Abbrechen' }
    );
    if (!ok) return;
    setBusy(btns, true);
    try {
        const uploaded = await uploadCurrentBackup({
            folder: DEFAULT_FOLDER,
            force: true,
            checkConflict: false
        });
        refreshStatus();
        const orgN =
            uploaded &&
            uploaded.payload &&
            uploaded.payload.schooltoolSummary &&
            uploaded.payload.schooltoolSummary.schulOrganigrammNodeCount;
        toast(
            orgN > 0
                ? 'Schul-/App-Daten in SharePoint gesichert (inkl. Schul-Organigramm: ' +
                      orgN +
                      ' Knoten).'
                : 'Schul-/App-Daten in SharePoint gesichert.'
        );
    } catch (e) {
        if (e && e.code === 'IT_LIBRARY_MISSING') {
            goToSetup(e.message);
            return;
        }
        if (e && e.code === 'spo-push-conflict') {
            toast(
                (e.message || 'Konflikt') +
                    ' Bitte Backup-Abgleich öffnen und entscheiden.',
                { kind: 'warning' }
            );
            return;
        }
        toast(e && e.message ? e.message : String(e));
    } finally {
        setBusy(btns, false);
    }
}

async function runLoad(btns) {
    if (!isSpoNetworkSyncAllowed()) {
        toast(localFirstStatusMessage());
        return;
    }
    if (!isReady()) {
        goToSetup('Zuerst die IT-Sicherungsbibliothek einrichten.');
        return;
    }
    setBusy(btns, true);
    try {
        const preview = await pickAndDownloadSharePointBackup({
            folder: DEFAULT_FOLDER,
            apply: false,
            title: 'Von SharePoint einlesen'
        });
        if (!preview) {
            toast('Abgebrochen.');
            return;
        }
        const obj = preview.payload || {};
        const item = preview.item || {};
        const choice = await resolveSharePointImportChoice(obj, {
            auto: false,
            remoteLastModified: item.lastModifiedDateTime || ''
        });
        if (choice === 'skip') {
            toast('Keine Änderung – lokaler und SharePoint-Stand sind identisch.');
            refreshStatus();
            return;
        }
        if (choice === 'keep-local') {
            toast('Lokaler Stand unverändert.');
            refreshStatus();
            return;
        }
        if (choice === 'push-local') {
            await uploadCurrentBackup({ folder: DEFAULT_FOLDER, force: true, checkConflict: false });
            toast('Lokaler Stand nach SharePoint gesichert.');
            refreshStatus();
            return;
        }
        const syncMeta = preview.meta
            ? Object.assign({}, preview.meta, {
                  at: new Date().toISOString(),
                  direction: 'pull',
                  dirty: false
              })
            : null;
        if (choice && typeof choice === 'object' && choice.action === 'import-selective') {
            const applied = applySharePointSelectiveImportLocally(obj, choice.selectedIds || [], syncMeta, {
                reload: false
            });
            const n = applied && applied.writtenKeys ? applied.writtenKeys.length : 0;
            toast('Auswahl übernommen (' + n + ' Schlüssel).');
        } else {
            applySharePointBackupLocally(obj, syncMeta, { reload: false });
            toast('Backup vollständig übernommen.');
        }
        const reload = await confirmAsync('Seite jetzt neu laden?', {
            title: 'Neu laden',
            confirmLabel: 'Neu laden',
            cancelLabel: 'Später'
        });
        if (reload) window.location.reload();
        else refreshStatus();
    } catch (e) {
        if (e && e.code === 'IT_LIBRARY_MISSING') {
            goToSetup(e.message);
            return;
        }
        toast(e && e.message ? e.message : String(e));
    } finally {
        setBusy(btns, false);
    }
}

function bind() {
    const uploadBtns = Array.from(document.querySelectorAll('[data-ms365-spo-sync="upload"]'));
    const loadBtns = Array.from(document.querySelectorAll('[data-ms365-spo-sync="load"]'));
    const setupLinks = Array.from(document.querySelectorAll('[data-ms365-spo-sync="setup"]'));
    const allAction = uploadBtns.concat(loadBtns);

    uploadBtns.forEach(function (btn) {
        if (btn.dataset.spoBound === '1') return;
        btn.dataset.spoBound = '1';
        btn.addEventListener('click', function (ev) {
            ev.preventDefault();
            runUpload(allAction);
        });
    });
    loadBtns.forEach(function (btn) {
        if (btn.dataset.spoBound === '1') return;
        btn.dataset.spoBound = '1';
        btn.addEventListener('click', function (ev) {
            ev.preventDefault();
            runLoad(allAction);
        });
    });
    setupLinks.forEach(function (el) {
        if (el.dataset.spoBound === '1') return;
        el.dataset.spoBound = '1';
        if (el.tagName === 'A' && !el.getAttribute('href')) {
            el.setAttribute('href', setupPageHref());
        } else if (el.tagName === 'A') {
            el.setAttribute('href', setupPageHref());
        } else {
            el.addEventListener('click', function (ev) {
                ev.preventDefault();
                window.location.href = setupPageHref();
            });
        }
    });

    // Delegation für dynamisch erzeugte Kopfzeilen-Buttons
    if (!bind._delegated) {
        bind._delegated = true;
        document.addEventListener(
            'click',
            function (ev) {
                const t = ev.target;
                if (!t || typeof t.closest !== 'function') return;
                const upload = t.closest('[data-ms365-spo-sync="upload"]');
                if (upload) {
                    if (upload.dataset.spoBound === '1') return;
                    ev.preventDefault();
                    ev.stopPropagation();
                    const btns = Array.from(
                        document.querySelectorAll(
                            '[data-ms365-spo-sync="upload"], [data-ms365-spo-sync="load"]'
                        )
                    );
                    runUpload(btns);
                    return;
                }
                const load = t.closest('[data-ms365-spo-sync="load"]');
                if (load) {
                    if (load.dataset.spoBound === '1') return;
                    ev.preventDefault();
                    ev.stopPropagation();
                    const btns = Array.from(
                        document.querySelectorAll(
                            '[data-ms365-spo-sync="upload"], [data-ms365-spo-sync="load"]'
                        )
                    );
                    runLoad(btns);
                }
            },
            true
        );
    }

    applyLocalFirstSpoUi(document);
    refreshStatus();
    try {
        window.addEventListener('ms365-tenant-settings-changed', refreshStatus);
        window.addEventListener('ms365-spo-sync-status', refreshStatus);
    } catch {
        /* ignore */
    }

    try {
        autoSync.start();
    } catch {
        /* ignore */
    }
}

function boot() {
    bind();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();

window.ms365StammdatenSpoSyncUi = {
    refreshStatus: refreshStatus,
    isReady: isReady,
    setupPageHref: setupPageHref,
    bind: bind
};

window.addEventListener('ms365-auth-widget-ready', bind);
