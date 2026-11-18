/**
 * Abgleich-UI: Modal oder Vollseite mit ausklappbaren Schlüssellisten
 * und selektivem Import (Bereiche wählen + Vorschau).
 */
import {
    buildBackupCoverageReport,
    compareBackupPayloads,
    storageKeyLabel
} from './stammdaten-sharepoint-sync-logic.js';
import { getLocalBackupSnapshot } from './stammdaten-sharepoint-import-prompt.js';
import {
    buildSelectiveImportPlan,
    describeSelectiveImport,
    statusLabelDe
} from './stammdaten-sharepoint-selective-import.js';

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

function groupLabelDe(group) {
    if (group === 'stammdaten') return 'Stammdaten';
    if (group === 'werkzeuge') return 'Werkzeuge & Rechte';
    if (group === 'weitere') return 'Weitere Schlüssel';
    return group || 'Bereiche';
}

function collectSelectedSectionIds(root) {
    if (!root) return [];
    return Array.from(root.querySelectorAll('[data-bcmp-section]:checked')).map(function (el) {
        return el.getAttribute('data-bcmp-section');
    }).filter(Boolean);
}

function setSelectiveChecks(root, mode, plan) {
    if (!root || !plan) return;
    const sections = Array.isArray(plan.sections) ? plan.sections : [];
    const byId = {};
    sections.forEach(function (s) {
        byId[s.id] = s;
    });
    root.querySelectorAll('[data-bcmp-section]').forEach(function (el) {
        const id = el.getAttribute('data-bcmp-section');
        const s = byId[id];
        if (!s) return;
        if (mode === 'none') el.checked = false;
        else if (mode === 'all-diff') {
            el.checked = s.status === 'differs' || s.status === 'remote-only';
        } else if (mode === 'recommended') {
            el.checked = !!s.recommended;
        }
    });
    updateSelectiveSummary(root, plan);
}

function updateSelectiveSummary(root, plan) {
    const el = root && root.querySelector('[data-bcmp-sel-summary]');
    if (!el) return;
    const ids = collectSelectedSectionIds(root);
    if (!ids.length) {
        el.textContent = 'Keine Bereiche ausgewählt.';
        return;
    }
    const labels = (plan.sections || [])
        .filter(function (s) {
            return ids.indexOf(s.id) !== -1;
        })
        .map(function (s) {
            return s.label;
        });
    el.textContent =
        ids.length +
        ' Bereich' +
        (ids.length === 1 ? '' : 'e') +
        ' ausgewählt: ' +
        labels.slice(0, 6).join(', ') +
        (labels.length > 6 ? ' …' : '');
}

function renderSelectivePreview(section) {
    const localSamples = (section.local && section.local.samples) || [];
    const remoteSamples = (section.remote && section.remote.samples) || [];
    if (!localSamples.length && !remoteSamples.length && !(section.local && section.local.summaryText)) {
        return '';
    }
    return (
        '<details class="ms365-bcmp-sel-preview">' +
        '<summary>Vorschau</summary>' +
        '<div class="ms365-bcmp-sel-preview__grid">' +
        '<div><strong>Lokal</strong><p class="muted">' +
        escapeHtml((section.local && section.local.summaryText) || '–') +
        '</p>' +
        (localSamples.length
            ? '<ul class="ms365-bcmp-sel-samples"><li>' +
              localSamples.map(escapeHtml).join('</li><li>') +
              '</li></ul>'
            : '') +
        '</div>' +
        '<div><strong>SharePoint</strong><p class="muted">' +
        escapeHtml((section.remote && section.remote.summaryText) || '–') +
        '</p>' +
        (remoteSamples.length
            ? '<ul class="ms365-bcmp-sel-samples"><li>' +
              remoteSamples.map(escapeHtml).join('</li><li>') +
              '</li></ul>'
            : '') +
        '</div></div></details>'
    );
}

function renderSelectiveSection(plan) {
    const p = plan || {};
    const sections = Array.isArray(p.sections) ? p.sections : [];
    if (!sections.length) return '';

    const groups = [];
    const seen = {};
    sections.forEach(function (s) {
        const g = s.group || 'weitere';
        if (!seen[g]) {
            seen[g] = true;
            groups.push(g);
        }
    });

    const blocks = groups
        .map(function (g) {
            const items = sections.filter(function (s) {
                return (s.group || 'weitere') === g;
            });
            const rows = items
                .map(function (s) {
                    const disabled = s.status === 'same' || s.status === 'missing' || s.status === 'local-only';
                    const checked = !disabled && s.recommended;
                    const statusCls =
                        s.status === 'differs' || s.status === 'remote-only'
                            ? 'ms365-bcmp-sel-status--diff'
                            : s.status === 'same'
                              ? 'ms365-bcmp-sel-status--ok'
                              : 'ms365-bcmp-sel-status--muted';
                    return (
                        '<li class="ms365-bcmp-sel-item' +
                        (disabled ? ' ms365-bcmp-sel-item--disabled' : '') +
                        '">' +
                        '<label class="ms365-bcmp-sel-label">' +
                        '<input type="checkbox" data-bcmp-section="' +
                        escapeHtml(s.id) +
                        '"' +
                        (checked ? ' checked' : '') +
                        (disabled ? ' disabled' : '') +
                        '>' +
                        '<span class="ms365-bcmp-sel-main">' +
                        '<span class="ms365-bcmp-sel-title">' +
                        escapeHtml(s.label) +
                        '</span>' +
                        '<span class="ms365-bcmp-sel-hint muted">' +
                        escapeHtml(s.hint || '') +
                        '</span>' +
                        '<span class="ms365-bcmp-sel-counts">' +
                        'Lokal: ' +
                        escapeHtml((s.local && s.local.summaryText) || '–') +
                        ' · SharePoint: ' +
                        escapeHtml((s.remote && s.remote.summaryText) || '–') +
                        '</span>' +
                        '</span>' +
                        '<span class="ms365-bcmp-sel-status ' +
                        statusCls +
                        '">' +
                        escapeHtml(statusLabelDe(s.status)) +
                        (s.recommended ? ' · empfohlen' : '') +
                        '</span>' +
                        '</label>' +
                        renderSelectivePreview(s) +
                        '</li>'
                    );
                })
                .join('');
            return (
                '<div class="ms365-bcmp-sel-group">' +
                '<h4 class="ms365-bcmp-sel-group__title">' +
                escapeHtml(groupLabelDe(g)) +
                '</h4>' +
                '<ul class="ms365-bcmp-sel-list">' +
                rows +
                '</ul></div>'
            );
        })
        .join('');

    return (
        '<section class="ms365-bcmp-selective" aria-label="Selektiver Import" data-bcmp-selective>' +
        '<h3 class="ms365-bcmp-diff__title">Selektiver Import</h3>' +
        '<p class="muted ms365-bcmp-sel-intro">' +
        'Nur markierte Bereiche aus SharePoint übernehmen – der Rest bleibt lokal. ' +
        'Gesamt-Backup: ' +
        escapeHtml(newerHintDe(p.newerSide)) +
        (p.remoteExportedAt ? ' (SP-Export ' + escapeHtml(formatWhen(p.remoteExportedAt)) + ')' : '') +
        '.' +
        '</p>' +
        '<div class="ms365-bcmp-sel-toolbar" role="group" aria-label="Schnellauswahl">' +
        '<button type="button" class="btn alt" data-bcmp-sel="recommended">Empfohlen</button>' +
        '<button type="button" class="btn alt" data-bcmp-sel="all-diff">Nur Unterschiede</button>' +
        '<button type="button" class="btn" data-bcmp-sel="none">Keine</button>' +
        '</div>' +
        blocks +
        '<p class="ms365-bcmp-sel-summary muted" data-bcmp-sel-summary></p>' +
        '</section>'
    );
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
    const selectiveHtml = o.selectivePlan ? renderSelectiveSection(o.selectivePlan) : '';

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
        selectiveHtml +
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
        '<button type="button" class="btn btn-success" data-bcmp-action="import-selective">' +
        '<i class="bi bi-check2-square" aria-hidden="true"></i>Auswahl übernehmen</button>' +
        '<button type="button" class="btn alt" data-bcmp-action="import" title="Ersetzt alle lokalen App-Daten">' +
        'Alles übernehmen</button>' +
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
    const plan = options.selectivePlan || null;

    function onClick(ev) {
        const selBtn = ev.target.closest('[data-bcmp-sel]');
        if (selBtn && mount.contains(selBtn)) {
            setSelectiveChecks(mount, selBtn.getAttribute('data-bcmp-sel'), plan);
            return;
        }
        const btn = ev.target.closest('[data-bcmp-action]');
        if (!btn || btn.disabled) return;
        handleAction(btn.getAttribute('data-bcmp-action'));
    }

    function onChange(ev) {
        if (ev.target && ev.target.matches && ev.target.matches('[data-bcmp-section]')) {
            updateSelectiveSummary(mount, plan);
        }
    }

    function handleAction(action) {
        if (action === 'open-page') {
            const ok = stashCompareSession({
                compare: cmp,
                remotePayload: options.remotePayload,
                localPayload: options.localPayload,
                remoteLastModified: options.remoteLastModified,
                title: options.title,
                selectivePlan: plan,
                coverage: options.coverage
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
        if (action === 'import-selective') {
            const selectedIds = collectSelectedSectionIds(mount);
            if (!selectedIds.length) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert('Bitte mindestens einen Bereich auswählen.');
                } else {
                    window.alert('Bitte mindestens einen Bereich auswählen.');
                }
                return;
            }
            let previewLines = [];
            try {
                const desc = describeSelectiveImport(
                    options.localPayload || getLocalBackupSnapshot(),
                    options.remotePayload,
                    selectedIds
                );
                previewLines = desc.previewLines || [];
            } catch (e) {
                if (typeof window.ms365ToastOrAlert === 'function') {
                    window.ms365ToastOrAlert(e && e.message ? e.message : String(e));
                }
                return;
            }
            const msg =
                'Diese Bereiche aus SharePoint übernehmen?\n\n' +
                previewLines.join('\n') +
                '\n\nAlles andere bleibt lokal.';
            const confirmFn =
                typeof window.ms365AppDialogConfirm === 'function'
                    ? window.ms365AppDialogConfirm
                    : function (m) {
                          return Promise.resolve(window.confirm(m));
                      };
            Promise.resolve(
                confirmFn(msg, {
                    title: 'Auswahl übernehmen',
                    okText: 'Übernehmen',
                    cancelText: 'Abbrechen',
                    kind: 'info'
                })
            ).then(function (ok) {
                if (ok) onChoice({ action: 'import-selective', selectedIds: selectedIds });
            });
            return;
        }
        if (action === 'import') {
            const msg =
                'Wirklich das gesamte Backup von SharePoint übernehmen?\n\n' +
                'Damit werden alle lokalen App-Daten (ms365-*/webuntis-*) ersetzt.';
            const confirmFn =
                typeof window.ms365AppDialogConfirm === 'function'
                    ? window.ms365AppDialogConfirm
                    : function (m) {
                          return Promise.resolve(window.confirm(m));
                      };
            Promise.resolve(
                confirmFn(msg, {
                    title: 'Alles übernehmen',
                    okText: 'Alles ersetzen',
                    cancelText: 'Abbrechen',
                    kind: 'warning',
                    danger: true
                })
            )
                .then(function (ok) {
                    if (ok) onChoice('import');
                })
                .catch(function (err) {
                    const m = err && err.message ? err.message : String(err || 'Unbekannter Fehler');
                    if (typeof window.ms365ToastOrAlert === 'function') {
                        window.ms365ToastOrAlert('Bestätigung fehlgeschlagen: ' + m);
                    } else {
                        window.alert('Bestätigung fehlgeschlagen: ' + m);
                    }
                });
            return;
        }
        if (action === 'keep-local' || action === 'push-local') {
            onChoice(action);
        }
    }

    mount.addEventListener('click', onClick);
    mount.addEventListener('change', onChange);
    updateSelectiveSummary(mount, plan);
    return function unbind() {
        mount.removeEventListener('click', onClick);
        mount.removeEventListener('change', onChange);
    };
}

function ensureSelectivePlan(options, cmp) {
    if (options.selectivePlan) return options.selectivePlan;
    const local = options.localPayload || getLocalBackupSnapshot();
    return buildSelectiveImportPlan(local, options.remotePayload || null, {
        newerSide: cmp && cmp.newerSide
    });
}

function renderIntoMount(mount, options, pageMode) {
    const cmp = options.compare || {};
    const coverage =
        options.coverage ||
        buildBackupCoverageReport(
            options.localPayload || getLocalBackupSnapshot(),
            options.remotePayload || null,
            cmp
        );
    const selectivePlan = ensureSelectivePlan(options, cmp);
    options.selectivePlan = selectivePlan;
    options.coverage = coverage;
    mount.innerHTML =
        renderCompareBody({
            compare: cmp,
            title: options.title,
            remoteLastModified: options.remoteLastModified,
            coverage: coverage,
            selectivePlan: selectivePlan
        }) + renderActions(cmp, pageMode);
    return cmp;
}

/**
 * @param {{
 *   compare: object,
 *   remotePayload?: object,
 *   localPayload?: object,
 *   remoteLastModified?: string,
 *   title?: string,
 *   pageMode?: boolean,
 *   selectivePlan?: object,
 *   coverage?: object
 * }} opts
 * @returns {Promise<'import'|'keep-local'|'push-local'|{ action: 'import-selective', selectedIds: string[] }>}
 */
export function promptBackupCompareChoice(opts) {
    const options = opts || {};
    if (!options.localPayload) {
        options.localPayload = getLocalBackupSnapshot();
    }

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
            if (ev.key !== 'Escape') return;
            // Bestätigungsdialog liegt darüber – Escape dort abfangen, Abgleich nicht schließen
            const appDlg = document.getElementById('ms365AppDialog');
            if (appDlg && appDlg.classList.contains('open')) return;
            ev.preventDefault();
            finish('keep-local');
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
            remotePayload: data.remotePayload || null,
            localPayload: data.localPayload || null,
            selectivePlan: data.selectivePlan || null,
            coverage: data.coverage || null
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

function isSelectiveChoice(choice) {
    return !!(choice && typeof choice === 'object' && choice.action === 'import-selective');
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
            await ensureAuthReadyForCompare();
            const { uploadCurrentBackup } = await import('./stammdaten-sharepoint-sync-api.js');
            const { DEFAULT_FOLDER } = await import('./stammdaten-sharepoint-sync-logic.js');
            // Explizite Nutzerwahl → Konflikt-Check aus (lokaler Stand soll gewinnen).
            await uploadCurrentBackup({ folder: DEFAULT_FOLDER, force: true, checkConflict: false });
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

    if (isSelectiveChoice(choice)) {
        if (!session.remotePayload) {
            toast('Backup-Daten fehlen in dieser Sitzung – bitte erneut Von SharePoint starten.');
            return;
        }
        try {
            const { applySharePointSelectiveImportLocally } = await import(
                './stammdaten-sharepoint-pull-apply.js'
            );
            const applied = applySharePointSelectiveImportLocally(
                session.remotePayload,
                choice.selectedIds || [],
                null,
                { reload: false, localPayload: session.localPayload }
            );
            const n = applied && applied.writtenKeys ? applied.writtenKeys.length : 0;
            if (mount) {
                mount.innerHTML =
                    '<div class="tm-panel"><p>Auswahl übernommen (' +
                    n +
                    ' Speicher-Schlüssel).</p><p><a class="btn" href="../index.html">Dashboard</a> ' +
                    '<a class="btn alt" href="../tenant.html">Stammdaten</a></p></div>';
            }
            toast('Auswahl übernommen.');
            const reload =
                typeof window.ms365AppDialogConfirm === 'function'
                    ? await window.ms365AppDialogConfirm('Seite jetzt neu laden?', {
                          title: 'Neu laden',
                          okText: 'Neu laden',
                          cancelText: 'Später'
                      })
                    : window.confirm('Seite jetzt neu laden?');
            if (reload) window.location.reload();
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
            await applySharePointBackupLocally(session.remotePayload, null, { reload: true });
        } catch (e) {
            toast(e && e.message ? e.message : String(e));
        }
    }
}

async function ensureAuthReadyForCompare(timeoutMs) {
    const limit = timeoutMs == null ? 12000 : Math.max(0, Number(timeoutMs) || 0);
    if (typeof window.ms365AuthAcquireToken === 'function') return true;
    await new Promise(function (resolve) {
        let done = false;
        function finish() {
            if (done) return;
            done = true;
            try {
                window.removeEventListener('ms365-auth-widget-ready', onReady);
            } catch {
                /* ignore */
            }
            resolve();
        }
        function onReady() {
            finish();
        }
        window.addEventListener('ms365-auth-widget-ready', onReady);
        const started = Date.now();
        const tick = window.setInterval(function () {
            if (typeof window.ms365AuthAcquireToken === 'function' || Date.now() - started >= limit) {
                window.clearInterval(tick);
                finish();
            }
        }, 150);
    });
    return typeof window.ms365AuthAcquireToken === 'function';
}

function downloadLocalBackupFile() {
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.downloadBackup !== 'function') {
        throw new Error('Browser-Backup-Modul fehlt.');
    }
    bb.downloadBackup();
}

async function pushLocalStandForce(mount) {
    const toast =
        typeof window.ms365ToastOrAlert === 'function'
            ? window.ms365ToastOrAlert
            : function (m) {
                  window.alert(m);
              };
    const ok =
        typeof window.ms365AppDialogConfirm === 'function'
            ? await window.ms365AppDialogConfirm(
                  'Diesen Browser-Stand jetzt nach SharePoint schreiben?\n\nSharePoint wird mit Ihrem lokalen Stand überschrieben (datierte Kopie bleibt erhalten).',
                  {
                      title: 'Lokalen Stand sichern',
                      okText: 'Nach SharePoint schreiben',
                      cancelText: 'Abbrechen',
                      kind: 'warning'
                  }
              )
            : window.confirm('Diesen Browser-Stand nach SharePoint schreiben?');
    if (!ok) return;
    try {
        const ready = await ensureAuthReadyForCompare();
        if (!ready) {
            throw new Error('Bitte oben rechts anmelden, dann erneut versuchen.');
        }
        if (typeof window.ms365AuthIsLoggedIn === 'function' && !window.ms365AuthIsLoggedIn()) {
            throw new Error('Bitte oben rechts anmelden, dann erneut „Nach SharePoint schreiben“.');
        }
        const { uploadCurrentBackup } = await import('./stammdaten-sharepoint-sync-api.js');
        const { DEFAULT_FOLDER } = await import('./stammdaten-sharepoint-sync-logic.js');
        await uploadCurrentBackup({ folder: DEFAULT_FOLDER, force: true, checkConflict: false });
        if (mount) {
            mount.innerHTML =
                '<div class="tm-panel">' +
                '<p><strong>Lokaler Stand ist auf SharePoint gesichert.</strong></p>' +
                '<p class="muted">Ihre Browser-Daten bleiben unverändert.</p>' +
                '<p><a class="btn" href="../index.html">Dashboard</a> ' +
                '<button type="button" class="btn alt" data-bcmp-run>Trotzdem abgleichen</button></p>' +
                '</div>';
            const again = mount.querySelector('[data-bcmp-run]');
            if (again) {
                again.addEventListener('click', function () {
                    void runLiveCompareOnPage(mount);
                });
            }
        }
        toast('Lokaler Stand nach SharePoint gesichert.', { kind: 'success' });
    } catch (e) {
        toast(e && e.message ? e.message : String(e), { kind: 'warning' });
    }
}

function renderCompareLanding(mount) {
    mount.innerHTML =
        '<div class="tm-panel ms365-bcmp-landing">' +
        '<h2 class="ms365-bcmp-landing__title">Lokal mit IT-Bibliothek vergleichen</h2>' +
        '<p class="muted">Wenn Ihr <strong>Browser-Stand passt</strong>: zuerst sichern (JSON-Datei und/oder SharePoint). Abgleichen nur, wenn Sie Unterschiede prüfen wollen.</p>' +
        '<div class="ms365-bcmp-landing__actions">' +
        '<button type="button" class="btn btn-success" data-bcmp-push-local title="Schreibt den aktuellen Browser-Stand nach SharePoint">' +
        '<i class="bi bi-cloud-arrow-up" aria-hidden="true"></i>Diesen Stand nach SharePoint schreiben</button>' +
        '<button type="button" class="btn" data-bcmp-download-local title="JSON-Datei herunterladen – unabhängig von SharePoint">' +
        '<i class="bi bi-download" aria-hidden="true"></i>Lokal als Datei speichern</button>' +
        '</div>' +
        '<hr class="ms365-bcmp-landing__rule" />' +
        '<p class="muted">Wählen Sie eine Backup-Datei aus der IT-Bibliothek (Ordner <code>Backups</code>) und vergleichen Sie sie mit dem Browser-Stand.</p>' +
        '<div class="ms365-bcmp-landing__actions">' +
        '<button type="button" class="btn" data-bcmp-run><i class="bi bi-arrow-repeat" aria-hidden="true"></i>Jetzt abgleichen</button>' +
        '<a class="btn" href="../index.html">Dashboard</a>' +
        '<a class="btn alt" href="stammdaten-uebergabe.html">Stammdaten-Übergabe</a>' +
        '</div>' +
        '<p class="ms365-bcmp-footnote muted">Anmeldung oben rechts nötig für SharePoint. Die lokale JSON-Datei funktioniert auch ohne Anmeldung.</p>' +
        '</div>';
    const btn = mount.querySelector('[data-bcmp-run]');
    if (btn) {
        btn.addEventListener('click', function () {
            void runLiveCompareOnPage(mount);
        });
    }
    const pushBtn = mount.querySelector('[data-bcmp-push-local]');
    if (pushBtn) {
        pushBtn.addEventListener('click', function () {
            void pushLocalStandForce(mount);
        });
    }
    const dlBtn = mount.querySelector('[data-bcmp-download-local]');
    if (dlBtn) {
        dlBtn.addEventListener('click', function () {
            try {
                downloadLocalBackupFile();
                const toast =
                    typeof window.ms365ToastOrAlert === 'function'
                        ? window.ms365ToastOrAlert
                        : function (m) {
                              window.alert(m);
                          };
                toast('Backup-Datei wird heruntergeladen.', { kind: 'success' });
            } catch (e) {
                window.alert(e && e.message ? e.message : String(e));
            }
        });
    }
}

function mountPageCompare(mount, session) {
    const options = {
        compare: session.compare,
        remotePayload: session.remotePayload,
        localPayload: session.localPayload,
        remoteLastModified: session.remoteLastModified,
        title: session.title || 'Sicherung abgleichen',
        coverage: session.coverage,
        selectivePlan: session.selectivePlan
    };
    const cmp = renderIntoMount(mount, options, true);
    session.selectivePlan = options.selectivePlan;
    bindCompareActions(mount, cmp, true, options, function (choice) {
        const buttons = mount.querySelectorAll('[data-bcmp-action]');
        buttons.forEach(function (b) {
            b.disabled = true;
        });
        void runPageChoice(choice, session);
    });
}

/**
 * Lädt Remote-Backup (nach Dateiauswahl) und vergleicht mit lokalem Stand (ohne Import).
 * @param {{ folder?: string, skipPicker?: boolean }} [opts]
 * @returns {Promise<null|{ local: object, remote: object, compare: object, coverage: object, remoteLastModified?: string, remoteFileName?: string }>}
 */
export async function runLiveBackupCompare(opts) {
    const options = opts || {};
    const { DEFAULT_FOLDER } = await import('./stammdaten-sharepoint-sync-logic.js');
    const folder = String(options.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
    const local = getLocalBackupSnapshot();
    if (!local) {
        throw new Error('Lokales Backup konnte nicht erstellt werden (Browser-Backup-Modul).');
    }
    let downloaded;
    if (options.skipPicker) {
        const { downloadCurrentBackup } = await import('./stammdaten-sharepoint-sync-api.js');
        downloaded = await downloadCurrentBackup({ folder: folder, apply: false });
    } else {
        const { pickAndDownloadSharePointBackup } = await import('./stammdaten-sharepoint-backup-picker-ui.js');
        downloaded = await pickAndDownloadSharePointBackup({
            folder: folder,
            apply: false,
            title: 'Backup für Abgleich wählen'
        });
        if (!downloaded) return null;
    }
    const remote = downloaded && downloaded.payload ? downloaded.payload : null;
    if (!remote) {
        throw new Error('SharePoint-Backup konnte nicht geladen werden.');
    }
    const cmp = compareBackupPayloads(local, remote);
    const coverage = buildBackupCoverageReport(local, remote, cmp);
    const selectivePlan = buildSelectiveImportPlan(local, remote, { newerSide: cmp.newerSide });
    const remoteLastModified =
        downloaded && downloaded.item && downloaded.item.lastModifiedDateTime
            ? downloaded.item.lastModifiedDateTime
            : '';
    const remoteFileName =
        (downloaded && downloaded.meta && downloaded.meta.fileName) ||
        (downloaded && downloaded.item && downloaded.item.name) ||
        '';
    return {
        local: local,
        remote: remote,
        compare: cmp,
        coverage: coverage,
        selectivePlan: selectivePlan,
        remoteLastModified: remoteLastModified,
        remoteFileName: remoteFileName
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
        '<div class="tm-panel"><p class="muted"><i class="bi bi-hourglass-split" aria-hidden="true"></i> Anmeldung prüfen und Backup vergleichen …</p></div>';
    try {
        const ready = await ensureAuthReadyForCompare();
        if (!ready) {
            throw new Error(
                'Anmeldung noch nicht bereit. Bitte oben rechts anmelden und „Erneut versuchen“.'
            );
        }
        if (typeof window.ms365AuthIsLoggedIn === 'function' && !window.ms365AuthIsLoggedIn()) {
            throw new Error('Bitte oben rechts anmelden, dann erneut abgleichen.');
        }
        const result = await runLiveBackupCompare();
        if (!result) {
            renderCompareLanding(mount);
            toast('Abgebrochen.');
            return;
        }
        const fileHint = result.remoteFileName ? ' · ' + result.remoteFileName : '';
        const session = {
            compare: result.compare,
            remotePayload: result.remote,
            localPayload: result.local,
            remoteLastModified: result.remoteLastModified,
            title: 'Lokal vs. IT-Bibliothek' + fileHint,
            coverage: result.coverage,
            selectivePlan: result.selectivePlan
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
            '<p class="muted">Ihr lokaler Browser-Stand wurde nicht verändert. Sie können ihn trotzdem als Datei speichern oder nach SharePoint schreiben.</p>' +
            '<div class="ms365-bcmp-landing__actions">' +
            '<button type="button" class="btn btn-success" data-bcmp-push-local><i class="bi bi-cloud-arrow-up" aria-hidden="true"></i>Diesen Stand nach SharePoint schreiben</button>' +
            '<button type="button" class="btn" data-bcmp-download-local><i class="bi bi-download" aria-hidden="true"></i>Lokal als Datei speichern</button>' +
            '<button type="button" class="btn alt" data-bcmp-run>Erneut abgleichen</button>' +
            '<a class="btn alt" href="stammdaten-uebergabe.html#setup">IT-Bibliothek einrichten</a>' +
            '</div></div>';
        const retry = mount.querySelector('[data-bcmp-run]');
        if (retry) {
            retry.addEventListener('click', function () {
                void runLiveCompareOnPage(mount);
            });
        }
        const pushBtn = mount.querySelector('[data-bcmp-push-local]');
        if (pushBtn) {
            pushBtn.addEventListener('click', function () {
                void pushLocalStandForce(mount);
            });
        }
        const dlBtn = mount.querySelector('[data-bcmp-download-local]');
        if (dlBtn) {
            dlBtn.addEventListener('click', function () {
                try {
                    downloadLocalBackupFile();
                    toast('Backup-Datei wird heruntergeladen.', { kind: 'success' });
                } catch (err) {
                    toast(err && err.message ? err.message : String(err));
                }
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
            // Erst Auth abwarten – sonst „MSAL-Widget nicht verfügbar“.
            void (async function () {
                await ensureAuthReadyForCompare();
                void runLiveCompareOnPage(mount);
            })();
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
