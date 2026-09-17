/**
 * Buttons: Stammdaten ↔ SharePoint IT-Bibliothek (Upload / Laden).
 * data-ms365-spo-sync="upload" | "load" | "setup"
 * Optional: #stammdatenSpoSyncStatus für Kurzstatus.
 */
import { DEFAULT_FOLDER, IT_LIBRARY_TITLE } from './stammdaten-sharepoint-sync-logic.js';
import {
    isReady,
    setupPageHref,
    loadItMeta,
    loadLocalSyncMeta,
    uploadCurrentBackup,
    downloadCurrentBackup
} from './stammdaten-sharepoint-sync-api.js';

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
    const it = loadItMeta();
    const m = loadLocalSyncMeta();
    if (!isReady()) {
        el.textContent =
            'SharePoint-IT-Bibliothek noch nicht eingerichtet – Sichern/Einlesen öffnet die Ersteinrichtung.';
        return;
    }
    const bits = ['Bereit: „' + (it.listTitle || IT_LIBRARY_TITLE) + '“'];
    if (m && m.at) {
        bits.push(
            'Zuletzt gesichert: ' + String(m.at).replace('T', ' ').replace(/\.\d+Z$/, '')
        );
    }
    el.textContent = bits.join(' · ');
}

async function runUpload(btns) {
    if (!isReady()) {
        goToSetup('Zuerst die IT-Sicherungsbibliothek einrichten.');
        return;
    }
    const ok = await confirmAsync(
        'Aktuelle Stammdaten (Browser-Backup) jetzt in die SharePoint-IT-Bibliothek „' +
            (loadItMeta().listTitle || IT_LIBRARY_TITLE) +
            '“ schreiben?',
        { title: 'Nach SharePoint sichern', confirmLabel: 'Sichern', cancelLabel: 'Abbrechen' }
    );
    if (!ok) return;
    setBusy(btns, true);
    try {
        await uploadCurrentBackup({ folder: DEFAULT_FOLDER, keepDated: false });
        refreshStatus();
        toast('Stammdaten in SharePoint gesichert.');
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
    if (!isReady()) {
        goToSetup('Zuerst die IT-Sicherungsbibliothek einrichten.');
        return;
    }
    setBusy(btns, true);
    try {
        const preview = await downloadCurrentBackup({ folder: DEFAULT_FOLDER, apply: false });
        const obj = preview.payload || {};
        const summary =
            (obj.schoolName || obj.domain || preview.item.name || 'Backup') +
            (obj.exportedAt ? ' · ' + String(obj.exportedAt).replace('T', ' ').slice(0, 19) : '');
        const ok = await confirmAsync(
            'Backup aus SharePoint übernehmen und lokale Stammdaten ersetzen?\n\n' + summary,
            { title: 'Von SharePoint einlesen', confirmLabel: 'Übernehmen', cancelLabel: 'Abbrechen' }
        );
        if (!ok) return;
        const bb = window.ms365BrowserBackup;
        if (!bb || typeof bb.importPayload !== 'function') throw new Error('Backup-Modul fehlt.');
        bb.importPayload(obj);
        toast('Backup übernommen.');
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

    refreshStatus();
    try {
        window.addEventListener('ms365-tenant-settings-changed', refreshStatus);
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
    setupPageHref: setupPageHref
};
