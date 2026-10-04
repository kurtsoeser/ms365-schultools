/**
 * Abgleich-UI: Modal oder Vollseite mit ausklappbaren Schlüssellisten.
 */
import { storageKeyLabel } from './stammdaten-sharepoint-sync-logic.js';

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

function renderCompareBody(opts) {
    const o = opts || {};
    const cmp = o.compare || {};
    const title = o.title || 'Sicherung abgleichen';
    const remoteLm = o.remoteLastModified ? formatWhen(o.remoteLastModified) : '';

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
    mount.innerHTML =
        renderCompareBody({
            compare: cmp,
            title: options.title,
            remoteLastModified: options.remoteLastModified
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

/** Vollseite: Session lesen und UI starten. */
export function initBackupComparePage() {
    const mount = document.querySelector('[data-ms365-bcmp-page]');
    if (!mount) return;
    const session = readCompareSession();
    if (!session || !session.compare) {
        mount.innerHTML =
            '<div class="tm-panel"><p class="muted">Kein Abgleich in dieser Sitzung. Starten Sie <strong>Von SharePoint</strong> am Dashboard, nutzen Sie den Abgleich nach der Anmeldung oder öffnen Sie im Dialog <strong>In neuem Tab öffnen</strong>.</p>' +
            '<p><a class="btn" href="../index.html">Dashboard</a></p></div>';
        return;
    }

    const options = {
        compare: session.compare,
        remotePayload: session.remotePayload,
        remoteLastModified: session.remoteLastModified,
        title: session.title || 'Sicherung abgleichen'
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

export default {
    promptBackupCompareChoice,
    stashCompareSession,
    takeCompareSession,
    readCompareSession,
    initBackupComparePage
};
