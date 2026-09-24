/**
 * Schreibt die bereinigte Kursteam-Liste als Unterrichtsbelegung
 * in die zentralen App-Daten (app-data-v2, aktuelles Schuljahr).
 */
import {
    buildSnapshotFromTeamsData,
    summarizeBelegung
} from '../../shared/unterrichtsbelegung-logic.js';

const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

function getYearPrefixFromUi() {
    try {
        const el = document.getElementById('yearPrefix');
        if (el && el.value) return String(el.value).trim();
    } catch {
        /* ignore */
    }
    return normStr(ns.yearPrefix);
}

function normStr(v) {
    return String(v ?? '').trim();
}

/**
 * @param {{ quiet?: boolean }} [options]
 * @returns {{ ok: boolean, snapshot: object|null, summary: object, error?: string }}
 */
ns.syncUnterrichtsbelegungToApp = function syncUnterrichtsbelegungToApp(options) {
    const quiet = !!(options && options.quiet);
    const result = {
        ok: false,
        snapshot: null,
        summary: { rows: 0, classes: 0, teachers: 0, yearPrefix: '' }
    };

    if (!ns.teamsGenerated || !Array.isArray(ns.teamsData) || !ns.teamsData.length) {
        return result;
    }

    const snapshot = buildSnapshotFromTeamsData(ns.teamsData, {
        yearPrefix: getYearPrefixFromUi(),
        source: 'kursteams'
    });
    if (!snapshot) return result;

    result.snapshot = snapshot;
    result.summary = summarizeBelegung(snapshot);

    const api = window.ms365AppDataV2;
    if (!api || typeof api.setUnterrichtsbelegung !== 'function') {
        result.error = 'App-Daten nicht verfügbar.';
        if (!quiet && typeof ns.showToast === 'function') {
            ns.showToast('Unterrichtsbelegung konnte nicht gespeichert werden (App-Daten fehlen).');
        }
        return result;
    }

    try {
        api.setUnterrichtsbelegung(snapshot);
        result.ok = true;
        if (!quiet && typeof ns.showToast === 'function') {
            const s = result.summary;
            ns.showToast(
                'Unterrichtsbelegung in App gespeichert: ' +
                    s.rows +
                    ' Einträge · ' +
                    s.classes +
                    ' Klassen · ' +
                    s.teachers +
                    ' Lehrkräfte' +
                    (s.yearPrefix ? ' (' + s.yearPrefix + ')' : '') +
                    '.'
            );
        }
        if (typeof ns.updateUnterrichtsbelegungHint === 'function') {
            ns.updateUnterrichtsbelegungHint();
        }
    } catch (e) {
        result.error = e && e.message ? String(e.message) : String(e);
        if (!quiet && typeof ns.showToast === 'function') {
            ns.showToast('Unterrichtsbelegung speichern fehlgeschlagen: ' + result.error);
        }
    }
    return result;
};

ns.getUnterrichtsbelegungForExport = function getUnterrichtsbelegungForExport() {
    if (!ns.teamsGenerated || !Array.isArray(ns.teamsData) || !ns.teamsData.length) return null;
    return buildSnapshotFromTeamsData(ns.teamsData, {
        yearPrefix: getYearPrefixFromUi(),
        source: 'kursteams'
    });
};

ns.updateUnterrichtsbelegungHint = function updateUnterrichtsbelegungHint() {
    const el = document.getElementById('unterrichtsbelegungHint');
    if (!el) return;
    const api = window.ms365AppDataV2;
    let snap = null;
    try {
        if (api && typeof api.getUnterrichtsbelegung === 'function') {
            snap = api.getUnterrichtsbelegung();
        }
    } catch {
        snap = null;
    }
    if (!snap || !Array.isArray(snap.rows) || !snap.rows.length) {
        el.hidden = true;
        el.textContent = '';
        return;
    }
    const s = summarizeBelegung(snap);
    const when = snap.updatedAt
        ? (() => {
              try {
                  return new Date(snap.updatedAt).toLocaleString('de-AT');
              } catch {
                  return snap.updatedAt;
              }
          })()
        : '';
    el.hidden = false;
    el.textContent =
        'App-Unterrichtsbelegung: ' +
        s.rows +
        ' Einträge · ' +
        s.classes +
        ' Klassen' +
        (s.yearPrefix ? ' · ' + s.yearPrefix : '') +
        (when ? ' · Stand ' + when : '') +
        ' (wird im Browser-Backup mitexportiert).';
};

window.ms365KursteamUnterrichtsbelegung = {
    syncUnterrichtsbelegungToApp: ns.syncUnterrichtsbelegungToApp,
    getUnterrichtsbelegungForExport: ns.getUnterrichtsbelegungForExport,
    updateUnterrichtsbelegungHint: ns.updateUnterrichtsbelegungHint
};
