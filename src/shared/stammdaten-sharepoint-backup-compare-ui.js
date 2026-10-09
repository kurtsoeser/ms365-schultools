/**
 * Abgleich-UI: Modal oder Vollseite mit ausklappbaren Schlüssellisten.
 */
import {
    buildBackupCoverageReport,
    compareBackupPayloads,
    storageKeyLabel
} from './stammdaten-sharepoint-sync-logic.js';
import { getLocalBackupSnapshot } from './stammdaten-sharepoint-import-prompt.js';

const SESSION_KEY = 'ms365-spo-compare-session-v1';
const SESSION_MAX_AGE_MS = 45 * 60 * 1000;
const OVERLAY_ID = 'ms365BackupCompare';

function formatWhen(iso) {
    const s = String(iso || '').trim();
    if (!s) return '–';
    return s.replace('T', ' ').replace(/\.\d{3}Z$/, ' UTC').slice(0, 22);
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function newerHintDe(side) {
    if (side === 'remote') return 'SharePoint-Stand wirkt aktueller';
    if (side === 'local') return 'Lokaler Stand wirkt aktueller';
    if (side === 'equal') return 'Zeitstempel gleich – Inhalt prüfen';
    return 'Aktualität unklar';
}

function renderKeyList(keys) {
    const list = Array.isArray(keys) ? keys : [];
    if (!list.length) {
        return '<p class="ms365-bcmp-empty muted">Keine Einträge</p>';
    }
    const items = list
        .map(function (k) {
            const label = storageKeyLabel(k);
            const showLabel = label !== k;
            return (
                '<li class="ms365-bcmp-key">' +
                '<code class="ms365-bcmp-key__code">' +
                escapeHtml(k) +
                '</code>' +
                (showLabel ? '<span class="ms365-bcmp-key__label">' + escapeHtml(label) + '</span>' : '') +
                '</li>'
            );
        })
        .join('');
    return '<ul class="ms365-bcmp-keys">' + items + '</ul>';
}

function renderDetails(summary, count, keys, open) {
    const n = Number(count) || 0;
    if (!n) return '';
    return (
        '<details class="ms365-bcmp-details"' +
        (open ? ' open' : '') +
        '>' +
        '<summary class="ms365-bcmp-details__summary">' +
        escapeHtml(summary) +
        ' <span class="ms365-bcmp-badge">' +
        n +
        '</span></summary>' +
        '<div class="ms365-bcmp-details__body">' +
        renderKeyList(keys) +
        '</div></details>'
    );
}

function spotlightStatusDe(status) {
    const s = String(status || '');
    if (s === 'same') return 'Gleich';
    if (s === 'differs') return 'Unterschiedlich';
    if (s === 'local-only') return 'Nur lokal';
    if (s === 'remote-only') return 'Nur SharePoint';
    return 'Fehlt';
}

function renderCoverageSection(coverage) {
    const c = coverage || {};
    const spotlight = Array.isArray(c.spotlight) ? c.spotlight : [];
    const rows = Array.isArray(c.schooltoolRows) ? c.schooltoolRows : [];
    if (!spotlight.length && !rows.length) return '';

    let banner = '';
    if (c.readyToSync) {
        banner =
            '<p class="ms365-bcmp-sync-ok" role="status"><i class="bi bi-check-circle" aria-hidden="true"></i> ' +
            'Lokal und IT-Bibliothek sind inhaltlich gleich (Fingerabdruck).</p>';
    } else if (c.mismatchSchooltool || c.mismatchSpotlight) {
        banner =
            '<p class="ms365-bcmp-sync-warn" role="status"><i class="bi bi-exclamation-triangle" aria-hidden="true"></i> ' +
            'Es gibt Abweichungen – unten Stichprobe und Schlüssellisten prüfen.</p>';
    }

    const schoolRows = rows
        .map(function (r) {
            const cls = r.match ? 'ms365-bcmp-cov-row--ok' : 'ms365-bcmp-cov-row--diff';
            return (
                '<tr class="' +
                cls +
                '"><th scope="row">' +
                escapeHtml(r.label) +
                '</th><td>' +
                escapeHtml(r.local) +
                '</td><td>' +
                escapeHtml(r.remote) +
                '</td></tr>'
            );
        })
        .join('');

    const spotRows = spotlight
        .map(function (s) {
            const cls =
                s.status === 'same'
                    ? 'ms365-bcmp-cov-row--ok'
                    : s.status === 'missing'
                      ? 'ms365-bcmp-cov-row--muted'
                      : 'ms365-bcmp-cov-row--diff';
            return (
                '<tr class="' +
                cls +
                '"><th scope="row">' +
                escapeHtml(s.label) +
                '</th><td>' +
                escapeHtml(spotlightStatusDe(s.status)) +
                '</td><td><code class="ms365-bcmp-key__code">' +
                escapeHtml(s.key) +
                '</code></td></tr>'
            );
        })
        .join('');

    return (
        '<section class="ms365-bcmp-coverage" aria-label="Stichprobe wichtiger Bereiche">' +
        '<h3 class="ms365-bcmp-diff__title">Stichprobe Stammdaten &amp; Konfiguration</h3>' +
        banner +
        '<h4 class="ms365-bcmp-cov-sub">Zentrale Schuldaten (ms365-schooltool-data-v2)</h4>' +
        '<table class="ms365-bcmp-cov-table"><thead><tr><th scope="col">Feld</th><th scope="col">Lokal</th><th scope="col">SharePoint</th></tr></thead><tbody>' +
        schoolRows +
        '</tbody></table>' +
        '<h4 class="ms365-bcmp-cov-sub">Wichtige localStorage-Schlüssel</h4>' +
        '<table class="ms365-bcmp-cov-table"><thead><tr><th scope="col">Bereich</th><th scope="col">Status</th><th scope="col">Schlüssel</th></tr></thead><tbody>' +
        spotRows +
        '</tbody></table>' +
        '</section>'
    );
}

function renderCompareBody(opts) {
    const o = opts || {};
    const cmp = o.compare || {};
    const title = o.title || 'Sicherung abgleichen';
    const remoteLm = o.remoteLastModified ? formatWhen(o.remoteLastModified) : '';
    const coverageHtml = o.coverage ? renderCoverageSection(o.coverage) : '';

    const hintClass =
        cmp.newerSide === 'local'
            ? 'ms365-bcmp-hint ms365-bcmp-hint--local'
            : cmp.newerSide === 'remote'
              ? 'ms365-bcmp-hint ms365-bcmp-hint--remote'
              : 'ms365-bcmp-hint';

    return (
        '<header class="ms365-bcmp-head">' +
        '<h2 class="ms365-bcmp-title" id="ms365BcmpTitle">' +
        escapeHtml(title) +
        '</h2>' +
        '<p class="' +
        hintClass +
        '">' +
        escapeHtml(newerHintDe(cmp.newerSide)) +
        '</p>' +
        '</header>' +
        '<div class="ms365-bcmp-cards">' +
        '<article class="ms365-bcmp-card">' +
        '<h3 class="ms365-bcmp-card__title">Lokal (dieser Browser)</h3>' +
        '<dl class="ms365-bcmp-meta">' +
        '<div><dt>Export</dt><dd>' +
        escapeHtml(formatWhen(cmp.localExportedAt)) +
        '</dd></div>' +
        '<div><dt>Schule</dt><dd>' +
        escapeHtml(cmp.localSchool || '–') +
        '</dd></div>' +
        '<div><dt>Schlüssel</dt><dd>' +
        escapeHtml(String(cmp.localKeyCount != null ? cmp.localKeyCount : '–')) +
        '</dd></div>' +
        '</dl></article>' +
        '<article class="ms365-bcmp-card ms365-bcmp-card--remote">' +
        '<h3 class="ms365-bcmp-card__title">SharePoint</h3>' +
        '<dl class="ms365-bcmp-meta">' +
        '<div><dt>Export</dt><dd>' +
        escapeHtml(formatWhen(cmp.remoteExportedAt)) +
        '</dd></div>' +
        (remoteLm
            ? '<div><dt>Datei geändert</dt><dd>' + escapeHtml(remoteLm) + '</dd></div>'
            : '') +
        '<div><dt>Schule</dt><dd>' +
        escapeHtml(cmp.remoteSchool || '–') +
        '</dd></div>' +
        '<div><dt>Schlüssel</dt><dd>' +
        escapeHtml(String(cmp.remoteKeyCount != null ? cmp.remoteKeyCount : '–')) +
        '</dd></div>' +
        '</dl></article>' +
        '</div>' +
        '<section class="ms365-bcmp-diff" aria-label="Unterschiede">' +
        '<h3 class="ms365-bcmp-diff__title">Unterschiede (localStorage)</h3>' +
        renderDetails('Inhalt geändert', cmp.changedCount, cmp.changedKeys, true) +
        renderDetails('Nur lokal', cmp.onlyLocalCount, cmp.onlyLocalKeys, false) +
        renderDetails('Nur auf SharePoint', cmp.onlyRemoteCount, cmp.onlyRemoteKeys, false) +
        '</section>' +
        coverageHtml +
        '<p class="ms365-bcmp-footnote muted">Es werden nur App-Schlüssel (<code>ms365-*</code>, <code>webuntis-*</code>) verglichen – nicht die Microsoft-Anmeldung.</p>'
    );
}

function renderActions(cmp, pageMode) {
    const showPush = cmp && cmp.newerSide === 'local';
    const openPage =
        !pageMode &&
        '<button type="button" class="btn alt ms365-bcmp-open-page" data-bcmp-action="open-page">' +
        '<i class="bi bi-box-arrow-up-right" aria-hidden="true"></i>In neuem Tab öffnen</button>';
    return (
        '<div class="ms365-bcmp-actions modal-actions">' +
        (openPage || '') +
        '<button type="button" class="btn" data-bcmp-action="keep-local">Lokal behalten</button>' +
        (showPush
            ? '<button type="button" class="btn alt" data-bcmp-action="push-local">Lokal nach SharePoint</button>'
            : '') +
        '<button type="button" class="btn btn-success" data-bcmp-action="import">Von SharePoint übernehmen</button>' +
        '</div>'
    );
}

function ensureOverlay() {
    let root = document.getElementById(OVERLAY_ID);
    if (root) return root;
    root = document.createElement('div');
    root.id = OVERLAY_ID;
    root.className = 'modal-overlay ms365-backup-compare-overlay';
    root.setAttribute('aria-hidden', 'true');
    root.innerHTML =
        '<div class="modal-box ms365-backup-compare-box" role="dialog" aria-modal="true" aria-labelledby="ms365BcmpTitle" tabindex="-1">' +
        '<div class="ms365-bcmp-scroll" data-bcmp-mount></div>' +
        '</div>';
    document.body.appendChild(root);
    root.addEventListener('click', function (e) {
        if (e.target === root) {
            root.dispatchEvent(new CustomEvent('ms365-bcmp-cancel', { bubbles: true }));
        }
    });
    return root;
}

function bindCompareActions(mount, cmp, pageMode, options, onChoice) {
    function onClick(ev) {
        const btn = ev.target.closest('[data-bcmp-action]');
        if (!btn || btn.disabled) return;
        handleAction(btn.getAttribute('data-bcmp-action'));
    }

    function handleAction(action) {
        if (action === 'open-page') {
            const ok = stashCompareSession({
                compare: cmp,
                remotePayload: options.remotePayload,
                remoteLastModified: options.remoteLastModified,
                title: options.title
            });
            if (!ok) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert(
                        'Abgleich zu groß für einen Tab-Wechsel – bitte hier entscheiden.'
                    );
                }
                return;
            }
            const href =
                (/\/tools\//i.test(location.pathname) ? '' : 'tools/') + 'stammdaten-backup-abgleich.html';
            window.open(href, '_blank', 'noopener');
            return;
        }
        if (action === 'import' || action === 'keep-local' || action === 'push-local') {
            onChoice(action);
        }
    }

    mount.addEventListener('click', onClick);
    return function unbind() {
        mount.removeEventListener('click', onClick);
    };
}

function renderIntoMount(mount, options, pageMode) {
    const cmp = options.compare || {};
    const coverage =
        options.coverage ||
        buildBackupCoverageReport(
            getLocalBackupSnapshot(),
            options.remotePayload || null,
            cmp
        );
    mount.innerHTML =
        renderCompareBody({
            compare: cmp,
            title: options.title,
            remoteLastModified: options.remoteLastModified,
            coverage: coverage
        }) + renderActions(cmp, pageMode);
    return cmp;
}

/**
 * @param {{
 *   compare: object,
 *   remotePayload?: object,
 *   remoteLastModified?: string,
 *   title?: string,
 *   pageMode?: boolean
 * }} opts
 * @returns {Promise<'import'|'keep-local'|'push-local'>}
 */
export function promptBackupCompareChoice(opts) {
    const options = opts || {};

    return new Promise(function (resolve) {
        const overlay = ensureOverlay();
        const mount = overlay.querySelector('[data-bcmp-mount]');
        if (!mount) {
            resolve('keep-local');
            return;
        }

        const cmp = renderIntoMount(mount, options, false);
        let unbind = function () {};

        function finish(choice) {
            unbind();
            overlay.classList.remove('open');
            overlay.setAttribute('aria-hidden', 'true');
            document.removeEventListener('keydown', onKey, true);
            resolve(choice);
        }

        function onKey(ev) {
            if (ev.key === 'Escape') {
                ev.preventDefault();
                finish('keep-local');
            }
        }

        unbind = bindCompareActions(mount, cmp, false, options, finish);
        document.addEventListener('keydown', onKey, true);

        overlay.classList.add('open');
        overlay.setAttribute('aria-hidden', 'false');
        const box = overlay.querySelector('.ms365-backup-compare-box');
        if (box) box.focus();

        overlay.addEventListener(
            'ms365-bcmp-cancel',
            function () {
                finish('keep-local');
            },
            { once: true }
        );
    });
}

export function stashCompareSession(data) {
    try {
        const payload = {
            at: Date.now(),
            compare: data.compare,
            remoteLastModified: data.remoteLastModified || '',
            title: data.title || '',
            remotePayload: data.remotePayload || null
        };
        const raw = JSON.stringify(payload);
        if (raw.length > 4_500_000) return false;
        sessionStorage.setItem(SESSION_KEY, raw);
        return true;
    } catch {
        return false;
    }
}

export function takeCompareSession() {
    try {
        const raw = sessionStorage.getItem(SESSION_KEY);
        if (!raw) return null;
        sessionStorage.removeItem(SESSION_KEY);
        const data = JSON.parse(raw);
        if (!data || !data.at || Date.now() - Number(data.at) > SESSION_MAX_AGE_MS) return null;
        return data;
    } catch {
        return null;
    }
}

export function readCompareSession() {
    try {
        const raw = sessionStorage.getItem(SESSION_KEY);
        if (!raw) return null;
        const data = JSON.parse(raw);
        if (!data || !data.at || Date.now() - Number(data.at) > SESSION_MAX_AGE_MS) return null;
        return data;
    } catch {
        return null;
    }
}

async function runPageChoice(choice, session) {
    const mount = document.querySelector('[data-ms365-bcmp-page]');
    const toast =
        typeof window.ms365ToastOrAlert === 'function'
            ? window.ms365ToastOrAlert
            : function (m) {
                  window.alert(m);
              };

    if (choice === 'keep-local') {
        if (mount) {
            mount.innerHTML =
                '<div class="tm-panel"><p>Lokaler Stand wurde behalten.</p><p><a class="btn" href="../index.html">Dashboard</a></p></div>';
        }
        toast('Lokaler Stand unverändert.');
        return;
    }

    if (choice === 'push-local') {
        try {
            const { uploadCurrentBackup } = await import('./stammdaten-sharepoint-sync-api.js');
            const { DEFAULT_FOLDER } = await import('./stammdaten-sharepoint-sync-logic.js');
            await uploadCurrentBackup({ folder: DEFAULT_FOLDER });
            if (mount) {
                mount.innerHTML =
                    '<div class="tm-panel"><p>Lokaler Stand wurde nach SharePoint gesichert.</p><p><a class="btn" href="../index.html">Dashboard</a></p></div>';
            }
            toast('Nach SharePoint gesichert.');
        } catch (e) {
            toast(e && e.message ? e.message : String(e));
        }
        return;
    }

    if (choice === 'import') {
        if (!session.remotePayload) {
            toast('Backup-Daten fehlen in dieser Sitzung – bitte erneut Von SharePoint starten.');
            return;
        }
        try {
            const { applySharePointBackupLocally } = await import('./stammdaten-sharepoint-pull-apply.js');
            applySharePointBackupLocally(session.remotePayload, null, { reload: true });
        } catch (e) {
            toast(e && e.message ? e.message : String(e));
        }
    }
}

function renderCompareLanding(mount) {
    mount.innerHTML =
        '<div class="tm-panel ms365-bcmp-landing">' +
        '<h2 class="ms365-bcmp-landing__title">Lokal mit IT-Bibliothek vergleichen</h2>' +
        '<p class="muted">Lädt <code>Backups/ms365-stammdaten-aktuell.json</code> aus der verknüpften IT-Bibliothek und vergleicht sie mit dem aktuellen Browser-Stand (inkl. Stichprobe zu Stammdaten, Dashboard, Planern).</p>' +
        '<p class="muted">Vor dem Vergleich werden offene Stammdaten-Auto-Speicherungen übernommen.</p>' +
        '<div class="ms365-bcmp-landing__actions">' +
        '<button type="button" class="btn btn-success" data-bcmp-run><i class="bi bi-arrow-repeat" aria-hidden="true"></i>Jetzt abgleichen</button>' +
        '<a class="btn" href="../index.html">Dashboard</a>' +
        '<a class="btn alt" href="stammdaten-uebergabe.html">Stammdaten-Übergabe</a>' +
        '</div>' +
        '<p class="ms365-bcmp-footnote muted">Alternativ: Nach Anmeldung erscheint der Abgleich automatisch, oder im Dialog <strong>In neuem Tab öffnen</strong>.</p>' +
        '</div>';
    const btn = mount.querySelector('[data-bcmp-run]');
    if (btn) {
        btn.addEventListener('click', function () {
            void runLiveCompareOnPage(mount);
        });
    }
}

function mountPageCompare(mount, session) {
    const options = {
        compare: session.compare,
        remotePayload: session.remotePayload,
        remoteLastModified: session.remoteLastModified,
        title: session.title || 'Sicherung abgleichen',
        coverage: session.coverage
    };
    const cmp = renderIntoMount(mount, options, true);
    bindCompareActions(mount, cmp, true, options, function (choice) {
        const buttons = mount.querySelectorAll('[data-bcmp-action]');
        buttons.forEach(function (b) {
            b.disabled = true;
        });
        void runPageChoice(choice, session);
    });
}

/**
 * Lädt Remote-Backup und vergleicht mit lokalem Stand (ohne Import).
 * @returns {Promise<{ local: object, remote: object, compare: object, coverage: object, remoteLastModified?: string }>}
 */
export async function runLiveBackupCompare() {
    const { downloadCurrentBackup } = await import('./stammdaten-sharepoint-sync-api.js');
    const { DEFAULT_FOLDER } = await import('./stammdaten-sharepoint-sync-logic.js');
    const local = getLocalBackupSnapshot();
    if (!local) {
        throw new Error('Lokales Backup konnte nicht erstellt werden (Browser-Backup-Modul).');
    }
    const downloaded = await downloadCurrentBackup({ folder: DEFAULT_FOLDER, apply: false });
    const remote = downloaded && downloaded.payload ? downloaded.payload : null;
    if (!remote) {
        throw new Error('SharePoint-Backup konnte nicht geladen werden.');
    }
    const cmp = compareBackupPayloads(local, remote);
    const coverage = buildBackupCoverageReport(local, remote, cmp);
    const remoteLastModified =
        downloaded && downloaded.item && downloaded.item.lastModifiedDateTime
            ? downloaded.item.lastModifiedDateTime
            : '';
    return {
        local: local,
        remote: remote,
        compare: cmp,
        coverage: coverage,
        remoteLastModified: remoteLastModified
    };
}

async function runLiveCompareOnPage(mount) {
    const toast =
        typeof window.ms365ToastOrAlert === 'function'
            ? window.ms365ToastOrAlert
            : function (m) {
                  window.alert(m);
              };
    mount.innerHTML =
        '<div class="tm-panel"><p class="muted"><i class="bi bi-hourglass-split" aria-hidden="true"></i> Lade IT-Bibliothek und vergleiche …</p></div>';
    try {
        const result = await runLiveBackupCompare();
        const session = {
            compare: result.compare,
            remotePayload: result.remote,
            remoteLastModified: result.remoteLastModified,
            title: 'Lokal vs. IT-Bibliothek',
            coverage: result.coverage
        };
        stashCompareSession(session);
        mountPageCompare(mount, session);
        if (result.coverage && result.coverage.readyToSync) {
            toast('Lokal und SharePoint sind identisch.', { kind: 'success', title: 'Backup-Abgleich' });
        }
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        mount.innerHTML =
            '<div class="tm-panel ms365-bcmp-error">' +
            '<p class="ms365-bcmp-sync-warn">Abgleich fehlgeschlagen: ' +
            escapeHtml(msg) +
            '</p>' +
            '<p class="muted">IT-Bibliothek verknüpft? Angemeldet mit Graph-Rechten? Datei <code>Backups/ms365-stammdaten-aktuell.json</code> vorhanden?</p>' +
            '<div class="ms365-bcmp-landing__actions">' +
            '<button type="button" class="btn btn-success" data-bcmp-run>Erneut versuchen</button>' +
            '<a class="btn alt" href="stammdaten-uebergabe.html#setup">IT-Bibliothek einrichten</a>' +
            '</div></div>';
        const retry = mount.querySelector('[data-bcmp-run]');
        if (retry) {
            retry.addEventListener('click', function () {
                void runLiveCompareOnPage(mount);
            });
        }
    }
}

/** Vollseite: Session lesen und UI starten. */
export function initBackupComparePage() {
    const mount = document.querySelector('[data-ms365-bcmp-page]');
    if (!mount) return;

    if (typeof window !== 'undefined') {
        window.ms365RunBackupCompare = runLiveBackupCompare;
    }

    const session = readCompareSession();
    if (session && session.compare) {
        mountPageCompare(mount, session);
        return;
    }

    renderCompareLanding(mount);

    try {
        const params = new URLSearchParams(window.location.search || '');
        if (params.get('run') === '1' || params.get('run') === 'true') {
            void runLiveCompareOnPage(mount);
        }
    } catch {
        /* ignore */
    }
}

export default {
    promptBackupCompareChoice,
    stashCompareSession,
    takeCompareSession,
    readCompareSession,
    runLiveBackupCompare,
    initBackupComparePage
};
