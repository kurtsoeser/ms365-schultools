/**
 * UI für den Querschnitts-Gruppenabgleich (tools/datenhygiene.html + Teaser auf index.html).
 * @file
 */
import {
    buildHygieneTargets,
    loadHygieneScanCache,
    runHygieneScan,
    summarizeHygieneScan
} from './membership-hygiene.js';
import { resolveToolsHref } from './app-paths.js';

function escapeHtml(s) {
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatWhen(iso) {
    if (!iso) return '';
    try {
        const d = new Date(iso);
        if (isNaN(d.getTime())) return '';
        return d.toLocaleString('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return '';
    }
}

function statusLabel(status) {
    if (status === 'ok') return 'Konsistent';
    if (status === 'mismatch') return 'Abweichung';
    if (status === 'unmatched') return 'Kein Match';
    if (status === 'empty-list') return 'Leere Liste';
    return 'Unbekannt';
}

function paintSummaryPills(container, counts) {
    if (!container) return;
    const c = counts || {};
    container.replaceChildren();
    const pills = [
        { tone: 'ok', label: 'Konsistent', n: c.ok || 0 },
        { tone: 'warn', label: 'Abweichung', n: c.mismatch || 0 },
        { tone: 'muted', label: 'Kein Match', n: c.unmatched || 0 }
    ];
    pills.forEach(function (p) {
        if (!p.n && p.tone !== 'warn') return;
        const span = document.createElement('span');
        span.className = 'dash-hygiene-pill dash-hygiene-pill--' + p.tone;
        span.textContent = p.n + ' ' + p.label;
        container.appendChild(span);
    });
    if (!container.childNodes.length) {
        const span = document.createElement('span');
        span.className = 'dash-hygiene-pill dash-hygiene-pill--muted';
        span.textContent = 'Noch keine Gruppen zum Prüfen';
        container.appendChild(span);
    }
}

/**
 * Kompakter Teaser auf index.html (letzter Scan-Stand, Link zum Werkzeug).
 */
export function mountHygieneTeaser() {
    const elSummary = document.getElementById('dashHygieneTeaserSummary');
    const elMeta = document.getElementById('dashHygieneTeaserMeta');
    if (!elSummary) return null;

    function loadContainer() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                return window.ms365AppDataV2.getContainer();
            }
        } catch {
            /* ignore */
        }
        return null;
    }

    function loadSettings() {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            return window.ms365TenantSettingsLoad();
        }
        return null;
    }

    function render() {
        const cached = loadHygieneScanCache();
        if (cached && cached.counts) {
            paintSummaryPills(elSummary, cached.counts);
            if (elMeta) {
                const when = formatWhen(cached.scannedAt || cached.savedAt);
                elMeta.textContent = when
                    ? 'Letzter Abgleich: ' + when + ' – Details im Werkzeug.'
                    : 'Details und Prüfung im Werkzeug.';
            }
            return;
        }
        const summary = summarizeHygieneScan(buildHygieneTargets(loadContainer(), loadSettings()), {});
        paintSummaryPills(elSummary, summary.counts);
        if (elMeta) {
            elMeta.textContent = 'Noch nicht geprüft – im Werkzeug mit Microsoft 365 abgleichen.';
        }
    }

    render();
    return { refresh: render };
}

/**
 * @param {object} cfg
 * @param {string} cfg.rootId
 */
export function mountMembershipHygieneDashboard(cfg) {
    const rootId = (cfg && cfg.rootId) || 'dashHygiene';
    const root = document.getElementById(rootId);
    if (!root) return null;

    const elSummary = document.getElementById('dashHygieneSummary');
    const elMeta = document.getElementById('dashHygieneMeta');
    const elBody = document.getElementById('dashHygieneBody');
    const elEmpty = document.getElementById('dashHygieneEmpty');
    const elScan = document.getElementById('dashHygieneScan');
    const elStatus = document.getElementById('dashHygieneStatus');

    function toast(msg) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
        else window.alert(msg);
    }

    function loadContainer() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                return window.ms365AppDataV2.getContainer();
            }
        } catch {
            /* ignore */
        }
        return null;
    }

    function loadSettings() {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            return window.ms365TenantSettingsLoad();
        }
        return null;
    }

    function renderSummary(counts) {
        paintSummaryPills(elSummary, counts);
    }

    function renderRows(rows, scannedAt, fromCache) {
        if (!elBody) return;
        elBody.replaceChildren();
        const list = Array.isArray(rows) ? rows : [];
        const hasMatched = list.some(function (r) {
            return r && r.groupId;
        });

        if (elEmpty) elEmpty.hidden = !!list.length;
        if (!list.length) {
            if (elMeta) {
                elMeta.textContent = 'Legen Sie Sammelgruppen oder Klassengruppen in den Werkzeugen an und matchen Sie sie mit Microsoft 365.';
            }
            renderSummary({});
            return;
        }

        list.forEach(function (row) {
            const tr = document.createElement('tr');
            tr.className = 'dash-hygiene-row dash-hygiene-row--' + (row.status || 'unknown');

            const tdLabel = document.createElement('td');
            tdLabel.innerHTML =
                '<span class="dash-hygiene-label">' +
                escapeHtml(row.label || '–') +
                '</span>' +
                (row.category === 'klasse'
                    ? '<span class="muted dash-hygiene-sub">Klassengruppe</span>'
                    : '<span class="muted dash-hygiene-sub">Sammelgruppe</span>');
            tr.appendChild(tdLabel);

            const tdList = document.createElement('td');
            tdList.className = 'dash-hygiene-num';
            tdList.textContent = String(typeof row.listCount === 'number' ? row.listCount : 0);
            tr.appendChild(tdList);

            const tdGroup = document.createElement('td');
            tdGroup.className = 'dash-hygiene-num';
            if (!row.groupId) tdGroup.textContent = '–';
            else if (row.groupCount === null || row.groupCount === undefined) tdGroup.textContent = '…';
            else if (row.groupCount < 0) tdGroup.textContent = '?';
            else tdGroup.textContent = String(row.groupCount);
            tr.appendChild(tdGroup);

            const tdStatus = document.createElement('td');
            const badge = document.createElement('span');
            badge.className = 'dash-hygiene-badge dash-hygiene-badge--' + (row.status || 'unknown');
            const icon =
                row.status === 'ok'
                    ? 'bi-check-circle-fill'
                    : row.status === 'mismatch'
                      ? 'bi-exclamation-triangle-fill'
                      : row.status === 'unmatched'
                        ? 'bi-dash-circle'
                        : 'bi-question-circle';
            badge.innerHTML = '<i class="bi ' + icon + '" aria-hidden="true"></i> ' + statusLabel(row.status);
            tdStatus.appendChild(badge);
            tr.appendChild(tdStatus);

            const tdAction = document.createElement('td');
            tdAction.className = 'dash-hygiene-action';
            if (row.toolHref) {
                const a = document.createElement('a');
                a.className = 'btn btn-sm';
                a.href = resolveToolsHref(row.toolHref);
                a.textContent = row.status === 'mismatch' ? 'Abgleich öffnen' : 'Werkzeug';
                a.title = row.reviewHint || '';
                tdAction.appendChild(a);
            } else {
                tdAction.textContent = '–';
            }
            tr.appendChild(tdAction);
            elBody.appendChild(tr);
        });

        if (elMeta) {
            let meta = '';
            if (scannedAt) {
                meta =
                    (fromCache ? 'Letzter Stand: ' : 'Geprüft: ') +
                    formatWhen(scannedAt) +
                    (hasMatched ? '' : ' · Noch keine gematchten Gruppen');
            } else if (hasMatched) {
                meta = 'Gematchte Gruppen erkannt – „Jetzt prüfen“ lädt Mitgliederzahlen aus Microsoft 365.';
            } else {
                meta = 'Noch keine gematchten Gruppen – zuerst in den Werkzeugen matchen.';
            }
            elMeta.textContent = meta;
        }
    }

    function renderFromTargetsOnly() {
        const targets = buildHygieneTargets(loadContainer(), loadSettings());
        const summary = summarizeHygieneScan(targets, {});
        renderSummary(summary.counts);
        renderRows(summary.rows, null, false);
    }

    function renderFromCache(cache) {
        if (!cache) {
            renderFromTargetsOnly();
            return;
        }
        renderSummary(cache.counts || {});
        renderRows(cache.rows || [], cache.scannedAt || cache.savedAt, true);
    }

    async function scan() {
        if (!window.ms365GraphUnifiedGroups) {
            toast('Graph-Modul fehlt – Seite neu laden.');
            return;
        }
        const G = window.ms365GraphUnifiedGroups;
        if (elScan) elScan.disabled = true;
        if (elStatus) elStatus.textContent = 'Prüfe Microsoft 365 …';
        try {
            const payload = await runHygieneScan({
                loadContainer: loadContainer,
                loadSettings: loadSettings,
                getGraphToken: function () {
                    return G.getGraphToken();
                },
                fetchGroupMemberCount: function (token, gid) {
                    return G.fetchGroupMemberCount(token, gid);
                }
            });
            renderSummary(payload.counts || {});
            renderRows(payload.rows || [], payload.scannedAt, false);
            if (elStatus) elStatus.textContent = 'Abgleich abgeschlossen.';
        } catch (e) {
            if (elStatus) elStatus.textContent = '';
            toast('Prüfung fehlgeschlagen: ' + (e && e.message ? e.message : e));
        } finally {
            if (elScan) elScan.disabled = false;
        }
    }

    if (elScan && !elScan.dataset.bound) {
        elScan.dataset.bound = '1';
        elScan.addEventListener('click', function () {
            void scan();
        });
    }

    const cached = loadHygieneScanCache();
    if (cached && Array.isArray(cached.rows) && cached.rows.length) {
        renderFromCache(cached);
    } else {
        renderFromTargetsOnly();
    }

    return { refresh: renderFromTargetsOnly, scan: scan };
}

const api = {
    mountMembershipHygieneDashboard: mountMembershipHygieneDashboard,
    mountHygieneTeaser: mountHygieneTeaser
};

if (typeof window !== 'undefined') {
    window.ms365MembershipHygieneDashboard = api;
}

export default api;
