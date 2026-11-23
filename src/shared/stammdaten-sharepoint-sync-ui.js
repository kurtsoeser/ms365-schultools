/**
 * Buttons: Stammdaten ↔ SharePoint IT-Bibliothek (manuell Upload / Laden / Setup).
 * data-ms365-spo-sync="upload" | "load" | "setup"
 *
 * Phase 4 / Lokal-First: nur manuell, kein Auto-Abgleich.
 * Auto-Sync-Modul startet als No-Op, solange isSpoNetworkSyncAllowed() false ist.
 */
import {
    isSpoManualBackupAllowed,
    isSpoNetworkSyncAllowed,
    localFirstStatusMessage,
    applyLocalFirstSpoUi,
    applyLocalFirstCompareUi
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
import { applySharePointBackupLocally } from './stammdaten-sharepoint-pull-apply.js';
import { pickAndDownloadSharePointBackup } from './stammdaten-sharepoint-backup-picker-ui.js';

function toast(m, opts) {
    if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(m, opts || {});
        return;
    }
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
    if (!isSpoManualBackupAllowed()) {
        el.textContent = localFirstStatusMessage();
        return;
    }
    const it = loadItMeta();
    const m = loadLocalSyncMeta();
    if (!isReady()) {
        el.textContent =
            'IT-Sicherungsbibliothek noch nicht eingerichtet – unter „Einrichten“ anlegen, danach manuell sichern.';
        return;
    }
    el.textContent = formatSyncStatusDe({
        ready: true,
        phase: 'manual',
        dirty: false,
        lastAt: m && m.at,
        lastDirection: m && m.direction,
        error: '',
        libraryTitle: (it && it.listTitle) || IT_LIBRARY_TITLE
    });
}

async function runUpload(btns) {
    if (!isSpoManualBackupAllowed()) {
        toast(localFirstStatusMessage());
        return;
    }
    if (!isReady()) {
        goToSetup('Zuerst die IT-Sicherungsbibliothek einrichten.');
        return;
    }
    const ok = await confirmAsync(
        'Aktuelles Browser-Backup jetzt in die SharePoint-IT-Bibliothek „' +
            (loadItMeta().listTitle || IT_LIBRARY_TITLE) +
            '“ schreiben?\n\n' +
            'Es wird die aktuelle Datei aktualisiert und zusätzlich eine datierte Kopie abgelegt ' +
            '(Versionierung). Es findet kein Abgleich mit dem Remote-Stand statt.',
        { title: 'In IT-Bibliothek sichern', confirmLabel: 'Sichern', cancelLabel: 'Abbrechen' }
    );
    if (!ok) return;
    setBusy(btns, true);
    try {
        const uploaded = await uploadCurrentBackup({
            folder: DEFAULT_FOLDER,
            keepDated: true,
            force: true,
            checkConflict: false
        });
        refreshStatus();
        const orgN =
            uploaded &&
            uploaded.payload &&
            uploaded.payload.schooltoolSummary &&
            uploaded.payload.schooltoolSummary.schulOrganigrammNodeCount;
        const dated =
            uploaded && uploaded.datedName
                ? ' Version: ' + uploaded.datedName + '.'
                : ' Mit datierter Kopie.';
        toast(
            orgN > 0
                ? 'Browser-Backup in IT-Bibliothek gesichert (inkl. Schul-Organigramm: ' +
                      orgN +
                      ' Knoten).' +
                      dated
                : 'Browser-Backup in IT-Bibliothek gesichert.' + dated
        );
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

async function runLoad(btns) {
    if (!isSpoManualBackupAllowed()) {
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
            title: 'Backup aus IT-Bibliothek laden'
        });
        if (!preview) {
            toast('Abgebrochen.');
            return;
        }
        const obj = preview.payload || {};
        const item = preview.item || {};
        const fileName = String((item && item.name) || preview.fileName || 'Backup').trim();
        const when = String((obj && obj.exportedAt) || (item && item.lastModifiedDateTime) || '').trim();
        const school = String((obj && (obj.schoolName || obj.domain)) || '').trim();
        const summaryRaw = String((obj && obj.inventorySummary) || '').trim();
        const summaryLines = summaryRaw
            ? summaryRaw
                  .split(/\s*[·•|]\s*|\s*;\s*/)
                  .map(function (s) {
                      return String(s || '').trim();
                  })
                  .filter(Boolean)
                  .map(function (s) {
                      return '• ' + s;
                  })
                  .join('\n')
            : '';
        const ok = await confirmAsync(
            'Dieses Backup jetzt in diesem Browser übernehmen?\n\n' +
                'Datei: ' +
                fileName +
                (when ? '\nStand: ' + when.replace('T', ' ').replace(/\.\d+Z$/, ' UTC') : '') +
                (school ? '\nSchule: ' + school : '') +
                (summaryLines ? '\n\nEnthalten:\n' + summaryLines : '') +
                '\n\nAlle lokalen Schuldaten und Werkzeugstände werden ersetzt. ' +
                'Microsoft-Anmeldung bleibt unberührt.',
            {
                title: 'Von IT-Bibliothek importieren',
                confirmLabel: 'Importieren',
                cancelLabel: 'Abbrechen',
                okText: 'Importieren',
                cancelText: 'Abbrechen',
                danger: true
            }
        );
        if (!ok) {
            toast('Import abgebrochen.');
            return;
        }
        const syncMeta = preview.meta
            ? Object.assign({}, preview.meta, {
                  at: new Date().toISOString(),
                  direction: 'pull',
                  dirty: false
              })
            : null;
        await applySharePointBackupLocally(obj, syncMeta, { reload: true });
        toast('Backup aus IT-Bibliothek übernommen. Die Seite wird neu geladen.');
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

function revealManualSpoUi(root) {
    const doc = root && root.querySelectorAll ? root : document;
    if (!doc || !doc.querySelectorAll) return;
    doc.querySelectorAll('[data-local-first-hide="spo-block"]').forEach(function (el) {
        const label = el.querySelector && el.querySelector('.ms365-backup-panel__label');
        const t = label ? String(label.textContent || '') : '';
        if (/Abgleich/i.test(t)) return;
        el.hidden = false;
        el.removeAttribute('aria-hidden');
    });
    const setup = doc.querySelector && doc.querySelector('[data-ms365-spo-sync="setup"]');
    if (setup) {
        setup.hidden = false;
        setup.removeAttribute('aria-hidden');
    }
    const note = doc.querySelector && doc.querySelector('[data-local-first-backup-note]');
    if (note) {
        note.hidden = true;
        note.textContent = '';
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
        if (el.tagName === 'A') {
            el.setAttribute('href', setupPageHref());
        } else {
            el.addEventListener('click', function (ev) {
                ev.preventDefault();
                window.location.href = setupPageHref();
            });
        }
    });

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
    applyLocalFirstCompareUi(document);
    if (isSpoManualBackupAllowed()) {
        revealManualSpoUi(document);
        const backup = document.getElementById('ms365BackupHeader');
        if (backup) {
            revealManualSpoUi(backup);
            applyLocalFirstCompareUi(backup);
        }
    }
    refreshStatus();
    try {
        window.addEventListener('ms365-tenant-settings-changed', refreshStatus);
        window.addEventListener('ms365-spo-sync-status', refreshStatus);
    } catch {
        /* ignore */
    }

    // Auto-Sync nur als No-Op starten (bleibt gesperrt), solange Network-Sync aus ist.
    try {
        if (typeof autoSync.start === 'function') autoSync.start();
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
    isManualAllowed: isSpoManualBackupAllowed,
    isNetworkSyncAllowed: isSpoNetworkSyncAllowed
};
