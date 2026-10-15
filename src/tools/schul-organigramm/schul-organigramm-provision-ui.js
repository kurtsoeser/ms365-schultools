/**
 * Abgleich-UI: Vorschau Landkarte ↔ M365, dann Graph + PowerShell/.cmd ausführen.
 */
import {
    collectSchulOrgStructureDiff,
    healSchulOrgStructureFromLists
} from './schul-organigramm-provision-diff.js';
import {
    buildSchulOrgM365AbgleichDraft,
    enrichSchulOrgM365AbgleichPlan,
    executeSchulOrgGraphAbgleich,
    buildAbgleichMailboxLauncher,
    countAbgleichActions,
    summarizeAbgleichCompare
} from './schul-organigramm-abgleich.js';
import { normalizeSchulOrganigramm } from './schul-organigramm-logic.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

async function acquireGraphToken() {
    const gug = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    let token = '';
    if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
        try {
            token = (await window.ms365AuthAcquireTokenSilent()) || '';
        } catch {
            token = '';
        }
    }
    if (!token && gug && typeof gug.getGraphToken === 'function') {
        try {
            token = (await gug.getGraphToken()) || '';
        } catch {
            token = '';
        }
    }
    if (!token && typeof window.ms365AuthAcquireToken === 'function') {
        token = (await window.ms365AuthAcquireToken()) || '';
    }
    return token;
}

function suggestedAdminUpn() {
    try {
        if (typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            const upn = String((info && (info.upn || info.username)) || '')
                .trim()
                .toLowerCase();
            if (upn.indexOf('@') !== -1) return upn;
        }
    } catch {
        /* ignore */
    }
    return '';
}

function downloadTextFile(filename, text, mime) {
    const blob = new Blob(['\uFEFF' + String(text || '')], {
        type: mime || 'text/plain;charset=utf-8'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}

function channelBadge(channel) {
    if (channel === 'powershell') {
        return '<span class="schul-org-diff__badge schul-org-diff__badge--ps">PowerShell</span>';
    }
    return '<span class="schul-org-diff__badge schul-org-diff__badge--graph">Graph</span>';
}

/**
 * @param {{
 *   loadSettings: () => object,
 *   saveSettings: (s: object) => void,
 *   onChanged?: () => void,
 *   toast?: (msg: string, kind?: string) => void
 * }} opts
 */
export function openSchulOrgProvisionDiffPanel(opts) {
    const o = opts || {};
    const externalToast = typeof o.toast === 'function' ? o.toast : null;
    const loadSettings = o.loadSettings;
    const saveSettings = o.saveSettings;
    if (typeof loadSettings !== 'function' || typeof saveSettings !== 'function') {
        if (externalToast) externalToast('Einstellungen nicht verfügbar.', 'warning');
        else if (typeof window.ms365Toast === 'function') {
            window.ms365Toast('Einstellungen nicht verfügbar.');
        }
        return;
    }

    const prev = document.querySelector('.schul-org-diff-overlay');
    if (prev && prev.parentNode) prev.parentNode.removeChild(prev);

    /** @type {ReturnType<typeof buildSchulOrgM365AbgleichDraft>|null} */
    let plan = null;
    let planEnriched = false;
    let busy = false;
    /** @type {string} */
    let statusMsg = '';
    /** @type {'info'|'success'|'warning'|''} */
    let statusKind = '';
    let previewLoading = false;
    /** @type {{ done: number, total: number, label: string }|null} */
    let previewProgress = null;

    function close() {
        const el = document.querySelector('.schul-org-diff-overlay');
        if (el && el.parentNode) el.parentNode.removeChild(el);
        document.removeEventListener('keydown', onEsc, true);
    }

    function onEsc(ev) {
        if (ev.key === 'Escape' && !busy) close();
    }

    /** Status im Panel – kein zweites Modal hinter dem Overlay. */
    function setPanelStatus(msg, kind) {
        statusMsg = String(msg || '');
        statusKind = kind || 'info';
        const el = document.querySelector('[data-diff-status]');
        if (!el) return;
        if (!statusMsg) {
            el.hidden = true;
            el.textContent = '';
            el.className = 'schul-org-diff__status';
            return;
        }
        el.hidden = false;
        el.textContent = statusMsg;
        el.className =
            'schul-org-diff__status' +
            (statusKind === 'success'
                ? ' schul-org-diff__status--ok'
                : statusKind === 'warning'
                  ? ' schul-org-diff__status--warn'
                  : '');
    }

    function structureDiff() {
        const settings = loadSettings();
        return collectSchulOrgStructureDiff(settings, settings.schulOrganigramm);
    }

    function ensureDraft() {
        if (!plan) {
            plan = buildSchulOrgM365AbgleichDraft(loadSettings());
            planEnriched = false;
        }
        return plan;
    }

    function syncSelectionFromDom(body) {
        if (!plan || !body) return;
        body.querySelectorAll('input[data-action-id]').forEach(function (inp) {
            const id = inp.getAttribute('data-action-id') || '';
            const a = plan.actions.find(function (x) {
                return x.id === id;
            });
            if (a) a.selected = /** @type {HTMLInputElement} */ (inp).checked;
        });
        plan.counts = countAbgleichActions(plan.actions);
    }

    function statusBadge(a) {
        if (a.status === 'noop') {
            return '<span class="schul-org-diff__state schul-org-diff__state--ok">passt</span>';
        }
        if (a.status === 'error') {
            return '<span class="schul-org-diff__state schul-org-diff__state--err">Fehler</span>';
        }
        if (a.status === 'needsGraph') {
            return '<span class="schul-org-diff__state schul-org-diff__state--wait">prüft …</span>';
        }
        if (a.type === 'createGroup') {
            return '<span class="schul-org-diff__state schul-org-diff__state--miss">Gruppe fehlt</span>';
        }
        if (a.type === 'createMailbox' || a.type === 'mailboxPermissions') {
            return '<span class="schul-org-diff__state schul-org-diff__state--miss">Postfach</span>';
        }
        return '<span class="schul-org-diff__state schul-org-diff__state--miss">fehlt in M365</span>';
    }

    function emailChips(emails, kind) {
        const list = Array.isArray(emails) ? emails : [];
        if (!list.length) return '';
        const cls =
            kind === 'ok'
                ? 'schul-org-diff__chip schul-org-diff__chip--ok'
                : kind === 'miss'
                  ? 'schul-org-diff__chip schul-org-diff__chip--miss'
                  : 'schul-org-diff__chip';
        const max = 8;
        const shown = list.slice(0, max);
        const more = list.length - shown.length;
        return (
            '<div class="schul-org-diff__chips">' +
            shown
                .map(function (e) {
                    return '<span class="' + cls + '">' + escapeHtml(e) + '</span>';
                })
                .join('') +
            (more > 0 ? '<span class="schul-org-diff__chip">+' + more + '</span>' : '') +
            '</div>'
        );
    }

    function renderActionRow(a) {
        const disabled = a.status === 'noop' || a.status === 'error';
        const checked = a.selected && !disabled;
        const already = Array.isArray(a.already) ? a.already : [];
        const toAdd = Array.isArray(a.toAdd) ? a.toAdd : [];
        const desired = Array.isArray(a.desired)
            ? a.desired
            : Array.isArray(a.memberEmails)
              ? a.memberEmails
              : [];
        let compareHtml = '';
        if (a.type === 'syncMembers' && planEnriched && a.status !== 'needsGraph') {
            compareHtml =
                '<div class="schul-org-diff__compare">' +
                '<div><span class="schul-org-diff__compare-label">Auf der Landkarte</span>' +
                (desired.length
                    ? emailChips(desired, '')
                    : '<span class="muted">keine Personen</span>') +
                '</div>' +
                '<div><span class="schul-org-diff__compare-label">Bereits in M365</span>' +
                (already.length
                    ? emailChips(already, 'ok')
                    : '<span class="muted">–</span>') +
                '</div>' +
                '<div><span class="schul-org-diff__compare-label">Fehlt in M365</span>' +
                (toAdd.length
                    ? emailChips(toAdd, 'miss')
                    : '<span class="muted">nichts</span>') +
                '</div></div>';
        } else if (a.type === 'createGroup') {
            compareHtml =
                '<div class="schul-org-diff__compare">' +
                '<div><span class="schul-org-diff__compare-label">Soll-Mitglieder (Landkarte)</span>' +
                (desired.length
                    ? emailChips(desired, 'miss')
                    : '<span class="muted">keine Personen</span>') +
                '</div></div>';
        }
        return (
            '<li class="' +
            (a.status === 'noop' ? 'is-noop' : a.status === 'error' ? 'is-err' : 'is-diff') +
            '">' +
            '<label>' +
            '<input type="checkbox" data-action-id="' +
            escapeHtml(a.id) +
            '"' +
            (checked ? ' checked' : '') +
            (disabled ? ' disabled' : '') +
            ' /> ' +
            '<span class="schul-org-diff__job-main">' +
            '<span class="schul-org-diff__job-top">' +
            channelBadge(a.channel) +
            ' ' +
            statusBadge(a) +
            ' <strong>' +
            escapeHtml(a.label) +
            '</strong>' +
            '<span class="muted"> · ' +
            escapeHtml(a.title) +
            '</span></span>' +
            '<span class="muted schul-org-diff__job-detail">' +
            escapeHtml(a.detail || '') +
            '</span>' +
            compareHtml +
            '</span></label></li>'
        );
    }

    function renderCompareSections(actions) {
        if (!actions.length) {
            return (
                '<p class="muted schul-org-diff__empty">' +
                'Keine M365-Verknüpfungen in der Landkarte. Pro Bereich unter „M365“ eine Gruppe wählen oder „Neu anlegen“ – dann erscheint der Vergleich.</p>'
            );
        }
        const cmp = summarizeAbgleichCompare(actions);
        function block(title, items, open) {
            if (!items.length) return '';
            return (
                '<details class="schul-org-diff__block"' +
                (open ? ' open' : '') +
                '>' +
                '<summary>' +
                escapeHtml(title) +
                ' <span class="schul-org-diff__badge">' +
                items.length +
                '</span></summary>' +
                '<ul class="schul-org-diff__jobs">' +
                items.map(renderActionRow).join('') +
                '</ul></details>'
            );
        }
        return (
            block('Fehlt / Abweichung (Mitglieder)', cmp.missing, true) +
            block('Gruppe noch nicht in M365', cmp.create, true) +
            block('Postfach / Rechte (PowerShell)', cmp.mailbox, true) +
            block('Bereits in Ordnung', cmp.ok, cmp.ok.length > 0 && !cmp.missing.length) +
            block('Fehler beim Prüfen', cmp.errors, true) +
            (cmp.pendingGraph.length
                ? block('Noch ohne Graph-Prüfung', cmp.pendingGraph, true)
                : '')
        );
    }

    function progressHtml() {
        if (!previewLoading && !previewProgress) return '';
        const done = previewProgress ? previewProgress.done : 0;
        const total = previewProgress ? previewProgress.total : 0;
        const pct = total > 0 ? Math.round((100 * done) / total) : previewLoading ? 8 : 0;
        const label = previewProgress
            ? 'Prüfe ' +
              done +
              '/' +
              total +
              (previewProgress.label ? ': ' + previewProgress.label : '')
            : 'Vorschau wird vorbereitet …';
        return (
            '<div class="schul-org-diff__progress" data-diff-progress>' +
            '<div class="schul-org-diff__progress-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="' +
            pct +
            '"><span style="width:' +
            pct +
            '%"></span></div>' +
            '<p class="schul-org-diff__progress-label">' +
            escapeHtml(label) +
            '</p></div>'
        );
    }

    function updateProgressDom() {
        const host = document.querySelector('[data-diff-progress-host]');
        if (!host) return;
        host.innerHTML = progressHtml();
        setPanelStatus(statusMsg, statusKind);
    }

    function paint() {
        const overlay = document.querySelector('.schul-org-diff-overlay');
        if (!overlay) return;
        const body = overlay.querySelector('[data-diff-body]');
        if (!body) return;
        const draft = ensureDraft();
        const c = draft.counts || countAbgleichActions(draft.actions);
        const cmp = summarizeAbgleichCompare(draft.actions);
        const sd = structureDiff();

        body.innerHTML =
            '<div class="schul-org-diff__stats">' +
            '<span class="schul-org-diff__stat' +
            (cmp.counts.ok ? ' is-ok' : '') +
            '"><strong>' +
            cmp.counts.ok +
            '</strong> passt</span>' +
            '<span class="schul-org-diff__stat' +
            (cmp.counts.missing || cmp.counts.create ? ' is-warn' : '') +
            '"><strong>' +
            (cmp.counts.missing + cmp.counts.create) +
            '</strong> fehlt / neu</span>' +
            '<span class="schul-org-diff__stat' +
            (cmp.counts.mailbox ? ' is-warn' : '') +
            '"><strong>' +
            cmp.counts.mailbox +
            '</strong> Postfach</span>' +
            '<span class="schul-org-diff__stat' +
            (cmp.counts.errors ? ' is-err' : '') +
            '"><strong>' +
            cmp.counts.errors +
            '</strong> Fehler</span>' +
            '</div>' +
            '<p class="schul-org-diff__status' +
            (statusKind === 'success'
                ? ' schul-org-diff__status--ok'
                : statusKind === 'warning'
                  ? ' schul-org-diff__status--warn'
                  : '') +
            '" data-diff-status' +
            (statusMsg ? '' : ' hidden') +
            '>' +
            escapeHtml(statusMsg) +
            '</p>' +
            '<div data-diff-progress-host>' +
            progressHtml() +
            '</div>' +
            '<section class="schul-org-diff__section">' +
            '<h3>Vergleich: Datenlandkarte ↔ Microsoft 365</h3>' +
            '<p class="muted">Graph prüft verknüpfte Gruppenmitglieder. Grün = schon vorhanden, Orange = fehlt noch in M365. Anschließend ausgewählte Aktionen ausführen.</p>' +
            '<div class="schul-org-diff__actions">' +
            '<button type="button" class="btn btn-primary" data-diff-preview' +
            (busy || previewLoading ? ' disabled' : '') +
            '>' +
            '<i class="bi bi-arrow-repeat" aria-hidden="true"></i> ' +
            (planEnriched ? 'Vorschau aktualisieren' : 'Vorschau laden') +
            '</button>' +
            '<button type="button" class="btn btn-success" data-diff-run' +
            (c.selected && !busy && !previewLoading ? '' : ' disabled') +
            '>' +
            '<i class="bi bi-play-fill" aria-hidden="true"></i> Ausgewählte durchführen</button>' +
            '</div>' +
            renderCompareSections(draft.actions) +
            '<div class="schul-org-diff__log" data-diff-log hidden></div>' +
            '</section>' +
            '<section class="schul-org-diff__section">' +
            '<h3>Struktur (Listen ↔ Landkarte)</h3>' +
            '<p class="muted">' +
            sd.counts.missingInLayout +
            ' fehlen auf der Landkarte · ' +
            sd.counts.orphanInLayout +
            ' Orphans</p>' +
            '<div class="schul-org-diff__actions">' +
            '<button type="button" class="btn btn-sm" data-diff-heal' +
            (sd.counts.missingInLayout && !busy ? '' : ' disabled') +
            '><i class="bi bi-diagram-3" aria-hidden="true"></i> Fehlende aus Listen ergänzen</button>' +
            '</div>' +
            '</section>';

        wireBody(body);
    }

    function appendLog(text, ok) {
        const log = document.querySelector('[data-diff-log]');
        if (!log) return;
        log.hidden = false;
        if (log.querySelector('.muted')) log.innerHTML = '';
        const p = document.createElement('p');
        p.textContent = text;
        p.className = ok ? 'schul-org-diff__ok' : 'schul-org-diff__err';
        log.appendChild(p);
    }

    function wireBody(body) {
        const previewBtn = body.querySelector('[data-diff-preview]');
        if (previewBtn) {
            previewBtn.addEventListener('click', function () {
                void runPreview();
            });
        }
        const runBtn = body.querySelector('[data-diff-run]');
        if (runBtn) {
            runBtn.addEventListener('click', function () {
                syncSelectionFromDom(body);
                void runExecute();
            });
        }
        body.querySelectorAll('input[data-action-id]').forEach(function (inp) {
            inp.addEventListener('change', function () {
                syncSelectionFromDom(body);
                const run = body.querySelector('[data-diff-run]');
                if (run && plan) {
                    /** @type {HTMLButtonElement} */ (run).disabled =
                        !plan.counts.selected || busy || previewLoading;
                }
            });
        });
        const healBtn = body.querySelector('[data-diff-heal]');
        if (healBtn) {
            healBtn.addEventListener('click', function () {
                const settings = loadSettings();
                const healed = healSchulOrgStructureFromLists(settings);
                settings.schulOrganigramm = healed.layout;
                saveSettings(settings);
                plan = null;
                planEnriched = false;
                setPanelStatus(
                    'Landkarte ergänzt (+' +
                        (healed.added || 0) +
                        (healed.healed ? ', bereinigt' : '') +
                        ').',
                    'success'
                );
                if (typeof o.onChanged === 'function') o.onChanged();
                paint();
                void runPreview();
            });
        }
    }

    async function runPreview() {
        if (busy) return;
        busy = true;
        previewLoading = true;
        previewProgress = null;
        planEnriched = false;
        setPanelStatus('Vorschau wird geladen …', 'info');
        paint();
        try {
            const settings = loadSettings();
            plan = buildSchulOrgM365AbgleichDraft(settings);
            paint();
            const gug = window.ms365GraphUnifiedGroups;
            const needsGraph = plan.actions.some(function (a) {
                return a.needsGraph;
            });
            if (!plan.actions.length) {
                planEnriched = true;
                setPanelStatus(
                    'Keine prüfbaren M365-Verknüpfungen. Bereiche in der Landkarte mit einer Gruppe verknüpfen oder „Neu anlegen“.',
                    'warning'
                );
            } else if (needsGraph) {
                if (!gug || typeof gug.fetchGroupMembers !== 'function') {
                    planEnriched = true;
                    setPanelStatus(
                        'Graph-Modul nicht geladen – nur lokaler Entwurf (ohne Live-Vergleich der Mitglieder).',
                        'warning'
                    );
                } else {
                    setPanelStatus('Anmeldung bei Microsoft Graph …', 'info');
                    updateProgressDom();
                    const token = await acquireGraphToken();
                    if (!token) throw new Error('Anmeldung für Microsoft Graph fehlgeschlagen.');
                    setPanelStatus('Vergleiche Gruppenmitglieder mit der Landkarte …', 'info');
                    updateProgressDom();
                    let setup = settings && settings.setup ? settings.setup : {};
                    try {
                        if (
                            window.ms365AppDataV2 &&
                            typeof window.ms365AppDataV2.getContainer === 'function'
                        ) {
                            const c = window.ms365AppDataV2.getContainer();
                            if (c && c.setup) setup = c.setup;
                        }
                    } catch {
                        /* ignore */
                    }
                    plan = await enrichSchulOrgM365AbgleichPlan(
                        plan,
                        {
                            token: token,
                            settings: settings,
                            setup: setup,
                            fetchGroupMembers: function (t, id) {
                                return gug.fetchGroupMembers(t, id);
                            },
                            fetchGroupOwners:
                                typeof gug.fetchGroupOwners === 'function'
                                    ? function (t, id) {
                                          return gug.fetchGroupOwners(t, id);
                                      }
                                    : undefined
                        },
                        function (ev) {
                            previewProgress = {
                                done: Number(ev && ev.done) || 0,
                                total: Number(ev && ev.total) || 0,
                                label: String((ev && ev.label) || '')
                            };
                            setPanelStatus(
                                'Prüfe Gruppe ' +
                                    previewProgress.done +
                                    '/' +
                                    previewProgress.total +
                                    (previewProgress.label
                                        ? ': ' + previewProgress.label
                                        : ''),
                                'info'
                            );
                            updateProgressDom();
                        }
                    );
                    planEnriched = true;
                    const cmp = plan.compare || summarizeAbgleichCompare(plan.actions);
                    setPanelStatus(
                        'Vergleich fertig: ' +
                            cmp.counts.ok +
                            ' passt · ' +
                            (cmp.counts.missing + cmp.counts.create) +
                            ' fehlt/neu · ' +
                            cmp.counts.mailbox +
                            ' Postfach' +
                            (cmp.counts.errors ? ' · ' + cmp.counts.errors + ' Fehler' : ''),
                        cmp.counts.errors ? 'warning' : 'success'
                    );
                }
            } else {
                planEnriched = true;
                const cmp = summarizeAbgleichCompare(plan.actions);
                setPanelStatus(
                    'Vorschau bereit (ohne Mitglieder-Live-Check): ' +
                        cmp.counts.create +
                        ' Gruppe(n) anlegen · ' +
                        cmp.counts.mailbox +
                        ' Postfach-Aktion(en).',
                    'success'
                );
            }
        } catch (e) {
            setPanelStatus(
                'Vorschau fehlgeschlagen: ' + (e && e.message ? e.message : e),
                'warning'
            );
        } finally {
            previewLoading = false;
            previewProgress = null;
            busy = false;
            paint();
        }
    }

    async function runExecute() {
        if (busy) return;
        ensureDraft();
        syncSelectionFromDom(document.querySelector('[data-diff-body]'));
        const selected = plan.actions.filter(function (a) {
            return a.selected && a.status !== 'noop' && a.status !== 'error';
        });
        if (!selected.length) {
            setPanelStatus('Keine Aktionen ausgewählt.', 'warning');
            paint();
            return;
        }
        const graphActions = selected.filter(function (a) {
            return a.channel === 'graph';
        });
        const psActions = selected.filter(function (a) {
            return a.channel === 'powershell';
        });

        busy = true;
        setPanelStatus('Abgleich läuft …', 'info');
        paint();
        const log = document.querySelector('[data-diff-log]');
        if (log) {
            log.hidden = false;
            log.innerHTML = '<p class="muted">Abgleich läuft …</p>';
        }
        let refreshPreview = false;
        try {
            const settings = loadSettings();
            if (graphActions.length) {
                const gug = window.ms365GraphUnifiedGroups;
                if (!gug || typeof gug.createGroupByKind !== 'function') {
                    throw new Error('Graph-Modul nicht geladen.');
                }
                const token = await acquireGraphToken();
                if (!token) throw new Error('Anmeldung für Microsoft Graph fehlgeschlagen.');
                const result = await executeSchulOrgGraphAbgleich(
                    settings,
                    graphActions,
                    {
                        token: token,
                        createGroupByKind: function (t, kind, name, nick, desc) {
                            return gug.createGroupByKind(t, kind, name, nick, desc);
                        },
                        ensureOwners:
                            typeof gug.ensureOwners === 'function'
                                ? function (t, gid, owners) {
                                      return gug.ensureOwners(t, gid, owners);
                                  }
                                : undefined,
                        syncEmailsToGroup:
                            typeof gug.syncEmailsToGroup === 'function'
                                ? function (t, gid, emails, label) {
                                      return gug.syncEmailsToGroup(t, gid, emails, label);
                                  }
                                : undefined,
                        sleep: typeof gug.sleep === 'function' ? gug.sleep : undefined
                    },
                    function (ev) {
                        appendLog(
                            (ev.ok ? '✓ ' : '✗ ') +
                                (ev.action && ev.action.title) +
                                ' · ' +
                                (ev.action && ev.action.label) +
                                (ev.message ? ' – ' + ev.message : ''),
                            !!ev.ok
                        );
                    }
                );
                settings.schulOrganigramm = normalizeSchulOrganigramm(result.layout);
                saveSettings(settings);
            }

            if (psActions.length) {
                const launcher = buildAbgleichMailboxLauncher(psActions, suggestedAdminUpn());
                if (launcher.cmd) {
                    downloadTextFile(
                        'ms365-abgleich-postfaecher.cmd',
                        launcher.cmd,
                        'application/x-bat;charset=utf-8'
                    );
                    appendLog(
                        '✓ PowerShell: ' +
                            launcher.count +
                            ' Postfach-Aktion(en) → .cmd heruntergeladen (Doppelklick, Admin-Konto wählen)',
                        true
                    );
                }
            }

            setPanelStatus(
                'Abgleich fertig: Graph ' +
                    graphActions.length +
                    ' · PowerShell ' +
                    psActions.length +
                    (psActions.length ? ' (.cmd gespeichert)' : ''),
                'success'
            );
            plan = null;
            planEnriched = false;
            if (typeof o.onChanged === 'function') o.onChanged();
            refreshPreview = true;
        } catch (e) {
            appendLog('Fehler: ' + (e && e.message ? e.message : e), false);
            setPanelStatus(
                'Abgleich fehlgeschlagen: ' + (e && e.message ? e.message : e),
                'warning'
            );
        } finally {
            busy = false;
            paint();
        }
        if (refreshPreview) void runPreview();
    }

    const overlay = document.createElement('div');
    overlay.className = 'schul-org-diff-overlay';
    overlay.innerHTML =
        '<div class="schul-org-diff-panel" role="dialog" aria-modal="true" aria-labelledby="schulOrgDiffTitle">' +
        '<header class="schul-org-diff-panel__head">' +
        '<div><h2 id="schulOrgDiffTitle">Abgleich mit Microsoft 365</h2>' +
        '<p class="muted">Vorschau lädt automatisch · Graph (Gruppen/Mitglieder) · PowerShell/.cmd (Postfächer)</p></div>' +
        '<button type="button" class="btn btn-sm" data-diff-close aria-label="Schließen">' +
        '<i class="bi bi-x-lg" aria-hidden="true"></i></button>' +
        '</header>' +
        '<div class="schul-org-diff-panel__body" data-diff-body></div>' +
        '</div>';
    document.body.appendChild(overlay);
    overlay.querySelector('[data-diff-close]').addEventListener('click', close);
    overlay.addEventListener('click', function (ev) {
        if (ev.target === overlay && !busy) close();
    });
    document.addEventListener('keydown', onEsc, true);
    paint();
    void runPreview();
}

export default { openSchulOrgProvisionDiffPanel };
